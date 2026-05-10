"""VUL (Variable UL) pricing engine.

UL with sub-account return as the credit. All other UL mechanics
unchanged (Phase 7, Section 3 of ``docs/seven_product_rollout_plan.md``).

Capabilities: ``supports_economic_scenario=True``,
``supports_monte_carlo=True``.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

import numpy as np

from annuity_model import pricing_projection as sp
from annuity_model._observability import traced
from annuity_model.account_value import AVConfig, evolve_account_value
from annuity_model.lapse import LapseAssumption
from annuity_model.mortality_2017_cso import MortalityTable2017CSO


@dataclass(frozen=True)
class VULContract:
    issue_age: int
    sex: Literal["male", "female"]
    smoker_class: Literal["nonsmoker", "smoker"] = "nonsmoker"
    face_amount: float = 250_000.0
    single_premium: float = 25_000.0
    premium_load_pct: float = 0.06
    monthly_expense_charge: float = 7.50
    db_type: Literal["return_of_av", "level_face"] = "level_face"
    horizon_age: int = 120
    payment_freq_per_year: int = 12
    subaccount_drift_annual: float = 0.06
    subaccount_vol_annual: float = 0.15

    @property
    def benefit_annual(self) -> float:
        return float(self.face_amount)


@dataclass(frozen=True)
class VULProjectionResult:
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
    expected_claim_cashflows: np.ndarray
    is_terminated_after_month: np.ndarray
    face_amount: float
    smoker_class: str


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


@traced("pricing.vul.deterministic")
def price_vul_single_premium(
    *,
    contract: VULContract,
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
) -> VULProjectionResult:
    if contract.payment_freq_per_year != 12:
        raise ValueError("VUL scaffold assumes monthly frequency.")
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
    # VUL credit: monthly sub-account simple return.
    monthly_credit = np.zeros(n_months, dtype=float)
    for t in range(1, n_months + 1):
        ratio = float(L[t] / L[t - 1]) if L[t - 1] > 0 else 1.0
        monthly_credit[t - 1] = ratio - 1.0

    av_cfg = AVConfig(
        initial_premium=float(contract.single_premium),
        premium_load_pct=float(contract.premium_load_pct),
        monthly_expense_charge=float(contract.monthly_expense_charge),
        db_type=contract.db_type,
        face_amount=float(contract.face_amount),
    )
    evol = evolve_account_value(
        config=av_cfg,
        n_months=n_months,
        monthly_credit_rate=monthly_credit,
        monthly_coi_q=monthly_q,
    )
    if np.any(evol.is_terminated_after_month):
        first_term = int(np.argmax(evol.is_terminated_after_month))
        if evol.is_terminated_after_month[first_term]:
            death_prob_month[first_term:] = 0.0
            survival_end[first_term:] = (
                float(survival_end[first_term - 1]) if first_term > 0 else 1.0
            )
            survival_start[first_term + 1 :] = float(survival_end[first_term])

    if lapse is not None:
        from annuity_model.lapse import combined_monthly_survival

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

    death_cf = evol.db_end_month * death_prob_month

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

    expected_benefit_cashflows = death_cf
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

    return VULProjectionResult(
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
        benefit_nominal_scheduled=np.array(evol.db_end_month, dtype=float),
        expense_nominal_scheduled=expense_sched,
        expense_annual_inflation=float(expense_annual_inflation),
        index_s0=float(s0),
        account_value_end_month=evol.account_value_end_month,
        db_end_month=evol.db_end_month,
        coi_dollars=evol.coi_dollars,
        nar_end_month=evol.nar_end_month,
        expected_claim_cashflows=expected_claim_cashflows,
        is_terminated_after_month=evol.is_terminated_after_month,
        face_amount=float(contract.face_amount),
        smoker_class=str(contract.smoker_class),
    )


@dataclass(frozen=True)
class VULMonteCarloResult:
    n_sims: int
    pv_benefit: np.ndarray
    pv_total_mean: float
    pv_benefit_mean: float
    av_end_mean: float


@traced("pricing.vul.monte_carlo")
def price_vul_single_premium_monte_carlo(
    *,
    contract: VULContract,
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
) -> VULMonteCarloResult:
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
    av_end = np.full(int(n_sims), np.nan, dtype=float)
    for i in range(int(n_sims)):
        path = idx_paths[i, :]
        levels_payment = path[1:].astype(float)
        res = price_vul_single_premium(
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
        av_end[i] = float(res.account_value_end_month[-1])
    return VULMonteCarloResult(
        n_sims=int(n_sims),
        pv_benefit=pvb,
        pv_total_mean=float(np.nanmean(pvb)),
        pv_benefit_mean=float(np.nanmean(pvb)),
        av_end_mean=float(np.nanmean(av_end)),
    )


def liability_path_from_vul_projection(pricing: VULProjectionResult) -> sp.LiabilityPath:
    return sp.LiabilityPath(
        times_years=np.asarray(pricing.times_years, dtype=float),
        expected_total_cashflows=np.asarray(pricing.expected_total_cashflows, dtype=float),
    )


from annuity_model.liability_dispatch import register_liability_path_converter  # noqa: E402

register_liability_path_converter("VULProjectionResult", liability_path_from_vul_projection)
