"""UL (Universal Life — single premium) pricing engine.

Single-premium UL with explicit COI. Monthly cycle:

    load -> declared rate credit -> COI -> flat expense charge

DB is Type A (``max(face, AV)``) — hardcoded in v1; UI does not expose
a selector. The dataclass field exists so v2 can add Type B without a
schema change.

Cashflow shape (Phase 5, Section 3 of ``docs/seven_product_rollout_plan.md``):

* **Death cashflow** at month ``t``: ``DB[t] * P(death in month t)``
  where ``DB[t] = max(face, AV[t])`` (Type A).
* **AV depletion** terminates the contract: cashflows are zero for
  every month after AV first reaches 0 (distinct from optional static
  lapse).

Capabilities: ``supports_economic_scenario=False``,
``supports_monte_carlo=False``.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

import numpy as np

import pricing_projection as sp
from _observability import traced
from account_value import AVConfig, evolve_account_value
from lapse import LapseAssumption
from mortality_2017_cso import MortalityTable2017CSO


@dataclass(frozen=True)
class ULContract:
    issue_age: int
    sex: Literal["male", "female"]
    smoker_class: Literal["nonsmoker", "smoker"] = "nonsmoker"
    face_amount: float = 250_000.0
    single_premium: float = 25_000.0
    premium_load_pct: float = 0.06
    monthly_expense_charge: float = 7.50
    declared_rate_annual: float = 0.04
    db_type: Literal["return_of_av", "level_face"] = "level_face"
    horizon_age: int = 120
    payment_freq_per_year: int = 12

    @property
    def benefit_annual(self) -> float:
        return float(self.face_amount)


@dataclass(frozen=True)
class ULProjectionResult:
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


def _ul_resolve_mortality(
    mortality, *, sex, smoker_class
) -> sp.MortalityTableQx | sp.MortalityTableRP2014MP2016:
    if mortality is None:
        return MortalityTable2017CSO.load(sex=sex, smoker_class=smoker_class).table
    if isinstance(mortality, MortalityTable2017CSO):
        return mortality.table
    return mortality


def _per_month_mortality_q(
    survival_end: np.ndarray,
) -> np.ndarray:
    n = survival_end.shape[0]
    surv_step = np.empty(n, dtype=float)
    surv_step[0] = survival_end[0]
    surv_step[1:] = survival_end[1:] / np.clip(survival_end[:-1], 1e-15, None)
    return np.clip(1.0 - surv_step, 0.0, 1.0)


@traced("pricing.ul.deterministic")
def price_ul_single_premium(
    *,
    contract: ULContract,
    yield_curve: sp.YieldCurve,
    mortality=None,
    horizon_age: int | None = None,
    spread: float = 0.0,
    valuation_year: int | None = None,
    expenses: sp.ExpenseAssumptions | None = None,
    expenses_csv_path: str = sp.DEFAULT_EXPENSES_CSV,
    index_scenario_csv_path: str | None = None,
    expense_annual_inflation: float = 0.0,
    lapse: LapseAssumption | None = None,
) -> ULProjectionResult:
    if contract.payment_freq_per_year != 12:
        raise ValueError("UL scaffold assumes monthly frequency.")
    if contract.face_amount <= 0:
        raise ValueError("face_amount must be > 0.")
    if contract.single_premium <= 0:
        raise ValueError("single_premium must be > 0.")
    del index_scenario_csv_path  # accepted for adapter signature parity

    mort = _ul_resolve_mortality(mortality, sex=contract.sex, smoker_class=contract.smoker_class)

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

    monthly_q = _per_month_mortality_q(survival_end)
    monthly_credit = np.full(
        n_months,
        (1.0 + float(contract.declared_rate_annual)) ** (1.0 / 12.0) - 1.0,
        dtype=float,
    )

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

    # AV-depletion handling: zero out cashflows after termination.
    # Survival is "held flat" by setting subsequent death-prob to 0.
    if np.any(evol.is_terminated_after_month):
        first_term = int(np.argmax(evol.is_terminated_after_month))
        if evol.is_terminated_after_month[first_term]:
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
        # death_prob_month overridden under combined decrements
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

    idx_level = np.full(n_months, 100.0, dtype=float)
    idx_zero = np.zeros(n_months, dtype=float)
    benefit_nominal = np.array(evol.db_end_month, dtype=float)

    return ULProjectionResult(
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
        index_level_at_payment=idx_level,
        index_simple_return=idx_zero.copy(),
        index_log_return=idx_zero.copy(),
        index_cumulative_return=idx_zero.copy(),
        benefit_nominal_scheduled=benefit_nominal,
        expense_nominal_scheduled=expense_sched,
        expense_annual_inflation=float(expense_annual_inflation),
        index_s0=100.0,
        account_value_end_month=evol.account_value_end_month,
        db_end_month=evol.db_end_month,
        coi_dollars=evol.coi_dollars,
        nar_end_month=evol.nar_end_month,
        expected_claim_cashflows=expected_claim_cashflows,
        is_terminated_after_month=evol.is_terminated_after_month,
        face_amount=float(contract.face_amount),
        smoker_class=str(contract.smoker_class),
    )


def liability_path_from_ul_projection(pricing: ULProjectionResult) -> sp.LiabilityPath:
    return sp.LiabilityPath(
        times_years=np.asarray(pricing.times_years, dtype=float),
        expected_total_cashflows=np.asarray(pricing.expected_total_cashflows, dtype=float),
    )


from liability_dispatch import register_liability_path_converter  # noqa: E402

register_liability_path_converter("ULProjectionResult", liability_path_from_ul_projection)
