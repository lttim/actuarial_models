"""IUL (Indexed UL) pricing engine.

UL with annual point-to-point crediting on segment anniversaries.
All other UL mechanics unchanged (Phase 6, Section 3 of
``docs/seven_product_rollout_plan.md``).

Capabilities: ``supports_economic_scenario=True``,
``supports_monte_carlo=True`` (reuses GBM index simulator).
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

import numpy as np

import pricing_projection as sp
from _observability import traced
from crediting import AnnualPointToPointCapped
from lapse import LapseAssumption
from mortality_2017_cso import MortalityTable2017CSO
from policy_features import (
    LevelPremiumSchedule,
    LoanTerms,
    MonthlySchedule,
    SurrenderChargeSchedule,
)


@dataclass(frozen=True)
class IULContract:
    issue_age: int
    sex: Literal["male", "female"]
    smoker_class: Literal["nonsmoker", "smoker"] = "nonsmoker"
    face_amount: float = 250_000.0
    single_premium: float = 25_000.0
    premium_load_pct: float = 0.06
    monthly_expense_charge: float = 7.50
    planned_premiums: LevelPremiumSchedule = LevelPremiumSchedule()
    withdrawals: MonthlySchedule = MonthlySchedule()
    loan_terms: LoanTerms = LoanTerms()
    surrender_charges: SurrenderChargeSchedule = SurrenderChargeSchedule()
    participation: float = 1.0
    cap: float = 0.10
    floor: float = 0.0
    db_type: Literal["return_of_av", "level_face"] = "level_face"
    horizon_age: int = 120
    segment_months: int = 12
    payment_freq_per_year: int = 12

    @property
    def benefit_annual(self) -> float:
        return float(self.face_amount)


@dataclass(frozen=True)
class IULProjectionResult:
    months: np.ndarray
    times_years: np.ndarray
    ages_at_payment: np.ndarray
    survival_to_payment: np.ndarray
    discount_factors: np.ndarray
    pv_benefit: float
    pv_monthly_expenses: float
    annuity_factor: float
    single_premium: float
    expected_benefit_cashflows: np.ndarray
    expected_expense_cashflows: np.ndarray
    expected_total_cashflows: np.ndarray
    reserve_times_years: np.ndarray
    economic_reserve: np.ndarray
    index_level_at_payment: np.ndarray
    index_simple_return: np.ndarray
    index_log_return: np.ndarray
    index_cumulative_return: np.ndarray
    benefit_nominal_scheduled: np.ndarray
    expense_nominal_scheduled: np.ndarray
    expense_annual_inflation: float
    index_s0: float
    account_value_end_month: np.ndarray
    db_end_month: np.ndarray
    coi_dollars: np.ndarray
    nar_end_month: np.ndarray
    segment_credited_rate: np.ndarray
    expected_claim_cashflows: np.ndarray
    is_terminated_after_month: np.ndarray
    face_amount: float
    smoker_class: str
    premium_cashflows: np.ndarray
    withdrawal_cashflows: np.ndarray
    loan_draw_cashflows: np.ndarray
    loan_repayment_cashflows: np.ndarray
    loan_balance_end_month: np.ndarray
    loan_interest_dollars: np.ndarray
    surrender_charge_dollars: np.ndarray
    surrender_value_end_month: np.ndarray
    net_death_benefit_end_month: np.ndarray


def _resolve_mortality(mortality, *, sex, smoker_class):
    if mortality is None:
        return MortalityTable2017CSO.load(sex=sex, smoker_class=smoker_class).table
    if isinstance(mortality, MortalityTable2017CSO):
        return mortality.table
    return mortality


def _per_month_q(survival_end: np.ndarray) -> np.ndarray:
    n = survival_end.shape[0]
    surv_step = np.empty(n, dtype=float)
    surv_step[0] = survival_end[0]
    surv_step[1:] = survival_end[1:] / np.clip(survival_end[:-1], 1e-15, None)
    return np.clip(1.0 - surv_step, 0.0, 1.0)


def _evolve_iul_policy_state(
    *,
    contract: IULContract,
    n_months: int,
    monthly_credit: np.ndarray,
    monthly_q: np.ndarray,
) -> tuple[
    np.ndarray,
    np.ndarray,
    np.ndarray,
    np.ndarray,
    np.ndarray,
    np.ndarray,
    np.ndarray,
    np.ndarray,
    np.ndarray,
    np.ndarray,
    np.ndarray,
    np.ndarray,
    np.ndarray,
    np.ndarray,
    np.ndarray,
]:
    """Monthly IUL state with scheduled premiums, withdrawals, and fixed loans."""

    cred = np.asarray(monthly_credit, dtype=float)
    qm = np.asarray(monthly_q, dtype=float)
    if cred.shape != (n_months,) or qm.shape != (n_months,):
        raise ValueError("monthly_credit and monthly_q must match n_months.")

    planned = contract.planned_premiums.values(n_months)
    premium = planned.copy()
    if n_months:
        premium[0] += float(contract.single_premium)
    withdrawals = contract.withdrawals.values(n_months)
    loan_draws = contract.loan_terms.draws.values(n_months)
    loan_repayments = contract.loan_terms.repayments.values(n_months)
    loan_rate_m = contract.loan_terms.monthly_rate()
    surrender_rates = contract.surrender_charges.monthly_rates(n_months)

    av_end = np.zeros(n_months, dtype=float)
    db_end = np.zeros(n_months, dtype=float)
    net_db_end = np.zeros(n_months, dtype=float)
    coi = np.zeros(n_months, dtype=float)
    nar = np.zeros(n_months, dtype=float)
    loan_bal = np.zeros(n_months, dtype=float)
    loan_interest = np.zeros(n_months, dtype=float)
    surrender_charge = np.zeros(n_months, dtype=float)
    surrender_value = np.zeros(n_months, dtype=float)
    credit_applied = np.zeros(n_months, dtype=float)
    terminated = np.zeros(n_months, dtype=bool)
    withdrawal_paid = np.zeros(n_months, dtype=float)

    av = 0.0
    loan = 0.0
    is_term = False
    for t in range(n_months):
        if is_term:
            terminated[t] = True
            loan *= 1.0 + loan_rate_m
            loan_bal[t] = loan
            continue
        li = loan * loan_rate_m
        loan += li
        loan_interest[t] = li

        prem = float(premium[t])
        av += prem * (1.0 - float(contract.premium_load_pct))
        cr = float(cred[t])
        av_after_credit = av * (1.0 + cr)
        gross_db = (
            max(float(contract.face_amount), av_after_credit)
            if contract.db_type == "level_face"
            else float(contract.face_amount) + max(0.0, av_after_credit)
        )
        nar_t = max(0.0, gross_db - av_after_credit)
        coi_t = float(qm[t]) * nar_t
        av = av_after_credit - coi_t - float(contract.monthly_expense_charge)

        draw = float(loan_draws[t])
        loan += draw
        repay = min(float(loan_repayments[t]), loan)
        loan -= repay
        loan_repayments[t] = repay

        wd = min(max(0.0, av), float(withdrawals[t]))
        av -= wd
        withdrawal_paid[t] = wd

        if av <= 0.0:
            av = 0.0
            is_term = True

        net_account = max(0.0, av - loan)
        sc = net_account * float(surrender_rates[t])
        av_end[t] = av
        db_end[t] = gross_db
        net_db_end[t] = max(0.0, gross_db - loan)
        coi[t] = coi_t
        nar[t] = nar_t
        loan_bal[t] = loan
        surrender_charge[t] = sc
        surrender_value[t] = max(0.0, net_account - sc)
        credit_applied[t] = cr
        terminated[t] = is_term

    return (
        av_end,
        db_end,
        net_db_end,
        coi,
        nar,
        credit_applied,
        terminated,
        premium,
        withdrawal_paid,
        loan_draws,
        loan_repayments,
        loan_bal,
        loan_interest,
        surrender_charge,
        surrender_value,
    )


@traced("pricing.iul.deterministic")
def price_iul_single_premium(
    *,
    contract: IULContract,
    yield_curve: sp.YieldCurve,
    mortality=None,
    horizon_age: int | None = None,
    spread: float = 0.0,
    valuation_year: int | None = None,
    expenses: sp.ExpenseAssumptions | None = None,
    expenses_csv_path: str = sp.DEFAULT_EXPENSES_CSV,
    index_scenario_csv_path: str | None = None,
    index_s0: float | None = None,
    index_levels_payment: np.ndarray | None = None,
    expense_annual_inflation: float = 0.0,
    lapse: LapseAssumption | None = None,
) -> IULProjectionResult:
    if contract.payment_freq_per_year != 12:
        raise ValueError("IUL scaffold assumes monthly frequency.")
    if contract.cap < contract.floor:
        raise ValueError("cap must be >= floor.")
    if contract.face_amount <= 0:
        raise ValueError("face_amount must be > 0.")
    if contract.single_premium <= 0:
        raise ValueError("single_premium must be > 0.")

    mort = _resolve_mortality(mortality, sex=contract.sex, smoker_class=contract.smoker_class)

    horizon = int(horizon_age) if horizon_age is not None else int(contract.horizon_age)
    dt = 1.0 / 12.0
    n_months = max(1, int(round((horizon - contract.issue_age) / dt)))

    months = np.arange(1, n_months + 1, dtype=int)
    times_years = months * dt
    ages_at_payment = contract.issue_age + times_years

    if valuation_year is None and isinstance(mort, sp.MortalityTableRP2014MP2016):
        raise ValueError("valuation_year must be provided when using MortalityTableRP2014MP2016.")
    survival_end = mort.monthly_survival_to_payment(
        issue_age=contract.issue_age,
        n_months=n_months,
        valuation_year=valuation_year,
    )
    survival_start = np.empty_like(survival_end)
    survival_start[0] = 1.0
    survival_start[1:] = survival_end[:-1]
    death_prob_month = np.clip(survival_start - survival_end, 0.0, 1.0)
    monthly_q = _per_month_q(survival_end)

    if index_levels_payment is not None:
        if index_scenario_csv_path is not None:
            raise ValueError(
                "Provide either index_scenario_csv_path or index_levels_payment, not both."
            )
        if index_s0 is None:
            raise ValueError("index_s0 must be provided when index_levels_payment is provided.")
        levels_payment = np.asarray(index_levels_payment, dtype=float)
        if levels_payment.shape != (n_months,):
            raise ValueError(f"index_levels_payment must have shape ({n_months},).")
        if (
            np.any(levels_payment <= 0.0)
            or np.any(~np.isfinite(levels_payment))
            or not np.isfinite(index_s0)
            or float(index_s0) <= 0.0
        ):
            raise ValueError("index levels and index_s0 must be finite and strictly positive.")
        s0 = float(index_s0)
    elif index_scenario_csv_path is None:
        s0, levels_payment = sp.flat_index_scenario(n_months)
    else:
        s0, levels_payment = sp.load_index_scenario_monthly_csv(
            index_scenario_csv_path, n_months=n_months
        )

    L = np.zeros(n_months + 1, dtype=float)
    L[0] = float(s0)
    L[1:] = np.asarray(levels_payment, dtype=float)

    seg = int(contract.segment_months)
    strategy = AnnualPointToPointCapped(
        participation=float(contract.participation),
        cap=float(contract.cap),
        floor=float(contract.floor),
    )
    monthly_credit = np.zeros(n_months, dtype=float)
    seg_credits = np.zeros(n_months, dtype=float)
    for m in range(1, n_months + 1):
        if m >= seg and (m % seg) == 0:
            raw = float(L[m] / L[m - seg] - 1.0)
            cr = strategy.credit_segment(raw_index_return=raw)
            monthly_credit[m - 1] = cr
            seg_credits[m - 1] = cr

    (
        av_end,
        db_end,
        net_db_end,
        coi_dollars,
        nar_end,
        credit_applied,
        terminated,
        premium_cashflows,
        withdrawal_paid,
        loan_draws,
        loan_repayments,
        loan_bal,
        loan_interest,
        surrender_charge,
        surrender_value,
    ) = _evolve_iul_policy_state(
        contract=contract,
        n_months=n_months,
        monthly_credit=monthly_credit,
        monthly_q=monthly_q,
    )
    if np.any(terminated):
        first_term = int(np.argmax(terminated))
        if terminated[first_term]:
            death_prob_month[first_term:] = 0.0
            survival_end[first_term:] = (
                float(survival_end[first_term - 1]) if first_term > 0 else 1.0
            )
            survival_start[first_term + 1 :] = float(survival_end[first_term])

    if lapse is not None:
        from lapse import combined_monthly_survival

        lapse_q_m = lapse.monthly_decrements(n_months)
        survival_combined = combined_monthly_survival(
            mortality_monthly_q=monthly_q,
            lapse_monthly_q=lapse_q_m,
        )
        survival_combined_start = np.empty_like(survival_combined)
        survival_combined_start[0] = 1.0
        survival_combined_start[1:] = survival_combined[:-1]
        death_prob_month = np.clip(survival_combined_start * monthly_q, 0.0, 1.0)
        survival_end = survival_combined
        survival_start = survival_combined_start

    death_cf = net_db_end * death_prob_month

    if expenses is None:
        try:
            expenses = sp.ExpenseAssumptions.load_from_csv(expenses_csv_path)
        except (FileNotFoundError, ValueError):
            expenses = sp.ExpenseAssumptions(0.0, 0.0, 0.0)
    g = (
        sp.monthly_rate_from_annual_inflation(float(expense_annual_inflation))
        if expense_annual_inflation
        else 0.0
    )
    expense_sched = float(expenses.monthly_expense_dollars) * (1.0 + g) ** np.arange(
        n_months, dtype=float
    )
    expected_expense_cashflows = expense_sched * survival_start

    expected_policy_access_cashflows = (withdrawal_paid + loan_draws) * survival_start
    expected_benefit_cashflows = death_cf + expected_policy_access_cashflows
    expected_total_cashflows = expected_benefit_cashflows + expected_expense_cashflows
    expected_claim_cashflows = death_cf

    df = yield_curve.discount_factors(times_years, spread=spread)
    pv_benefit = float(np.sum(expected_benefit_cashflows * df))
    pv_monthly_expenses = float(np.sum(expected_expense_cashflows * df))
    annuity_factor = float(np.sum(survival_end * df))

    reserve_times_years = np.concatenate(([0.0], times_years))
    economic_reserve = np.zeros(n_months + 1, dtype=float)
    pv_remaining = np.zeros(n_months + 1, dtype=float)
    for i in range(n_months - 1, -1, -1):
        pv_remaining[i] = float(expected_total_cashflows[i] * df[i] + pv_remaining[i + 1])
    economic_reserve[0] = float(pv_remaining[0])
    for i in range(n_months):
        if i + 1 >= n_months:
            economic_reserve[i + 1] = 0.0
            continue
        if survival_end[i] <= 0.0 or df[i] <= 0.0:
            economic_reserve[i + 1] = 0.0
            continue
        economic_reserve[i + 1] = float(pv_remaining[i + 1] / (survival_end[i] * df[i]))

    simp_ret = np.zeros(n_months, dtype=float)
    log_ret = np.zeros(n_months, dtype=float)
    cumu_ret = np.zeros(n_months, dtype=float)
    prev = float(s0)
    for k in range(n_months):
        cur = float(levels_payment[k])
        simp_ret[k] = cur / prev - 1.0 if prev > 0 else 0.0
        log_ret[k] = np.log(cur / prev) if cur > 0 and prev > 0 else float("nan")
        cumu_ret[k] = cur / float(s0) - 1.0
        prev = cur

    return IULProjectionResult(
        months=months,
        times_years=times_years,
        ages_at_payment=ages_at_payment,
        survival_to_payment=survival_end,
        discount_factors=df,
        pv_benefit=pv_benefit,
        pv_monthly_expenses=pv_monthly_expenses,
        annuity_factor=annuity_factor,
        single_premium=float(contract.single_premium),
        expected_benefit_cashflows=expected_benefit_cashflows,
        expected_expense_cashflows=expected_expense_cashflows,
        expected_total_cashflows=expected_total_cashflows,
        reserve_times_years=reserve_times_years,
        economic_reserve=economic_reserve,
        index_level_at_payment=levels_payment,
        index_simple_return=simp_ret,
        index_log_return=log_ret,
        index_cumulative_return=cumu_ret,
        benefit_nominal_scheduled=np.array(net_db_end, dtype=float),
        expense_nominal_scheduled=expense_sched,
        expense_annual_inflation=float(expense_annual_inflation),
        index_s0=float(s0),
        account_value_end_month=av_end,
        db_end_month=db_end,
        coi_dollars=coi_dollars,
        nar_end_month=nar_end,
        segment_credited_rate=seg_credits,
        expected_claim_cashflows=expected_claim_cashflows,
        is_terminated_after_month=terminated,
        face_amount=float(contract.face_amount),
        smoker_class=str(contract.smoker_class),
        premium_cashflows=premium_cashflows,
        withdrawal_cashflows=withdrawal_paid * survival_start,
        loan_draw_cashflows=loan_draws * survival_start,
        loan_repayment_cashflows=loan_repayments * survival_start,
        loan_balance_end_month=loan_bal,
        loan_interest_dollars=loan_interest,
        surrender_charge_dollars=surrender_charge,
        surrender_value_end_month=surrender_value,
        net_death_benefit_end_month=net_db_end,
    )


@dataclass(frozen=True)
class IULMonteCarloResult:
    n_sims: int
    pv_benefit: np.ndarray
    pv_total_mean: float
    pv_benefit_mean: float


@traced("pricing.iul.monte_carlo")
def price_iul_single_premium_monte_carlo(
    *,
    contract: IULContract,
    yield_curve: sp.YieldCurve,
    mortality=None,
    horizon_age: int | None = None,
    spread: float = 0.0,
    valuation_year: int | None = None,
    expenses: sp.ExpenseAssumptions | None = None,
    expenses_csv_path: str = sp.DEFAULT_EXPENSES_CSV,
    expense_annual_inflation: float = 0.0,
    n_sims: int = 200,
    annual_drift: float = 0.06,
    annual_vol: float = 0.15,
    seed: int | None = None,
    s0: float = 100.0,
) -> IULMonteCarloResult:
    horizon = int(horizon_age) if horizon_age is not None else int(contract.horizon_age)
    dt = 1.0 / 12.0
    n_months = max(1, int(round((horizon - contract.issue_age) / dt)))
    idx_paths = sp.simulate_index_levels_gbm(
        n_sims=n_sims,
        n_months=n_months,
        s0=s0,
        annual_drift=annual_drift,
        annual_vol=annual_vol,
        seed=seed,
    )
    pvb = np.full(int(n_sims), np.nan, dtype=float)
    for i in range(int(n_sims)):
        path = idx_paths[i, :]
        levels_payment = path[1:].astype(float)
        res = price_iul_single_premium(
            contract=contract,
            yield_curve=yield_curve,
            mortality=mortality,
            horizon_age=horizon,
            spread=spread,
            valuation_year=valuation_year,
            expenses=expenses,
            expenses_csv_path=expenses_csv_path,
            index_s0=float(path[0]),
            index_levels_payment=levels_payment,
            expense_annual_inflation=expense_annual_inflation,
        )
        pvb[i] = float(res.pv_benefit)
    return IULMonteCarloResult(
        n_sims=int(n_sims),
        pv_benefit=pvb,
        pv_total_mean=float(np.nanmean(pvb)),
        pv_benefit_mean=float(np.nanmean(pvb)),
    )


def liability_path_from_iul_projection(pricing: IULProjectionResult) -> sp.LiabilityPath:
    return sp.LiabilityPath(
        times_years=np.asarray(pricing.times_years, dtype=float),
        expected_total_cashflows=np.asarray(pricing.expected_total_cashflows, dtype=float),
    )


from liability_dispatch import register_liability_path_converter  # noqa: E402

register_liability_path_converter("IULProjectionResult", liability_path_from_iul_projection)
