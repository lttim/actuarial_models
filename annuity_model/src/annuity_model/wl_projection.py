"""WL (Whole Life — single premium) pricing engine.

Single-premium paid-up whole life. Level death benefit ``face_amount``
payable at month-end of death for life. Mortality is 2017 CSO Ultimate
(sex × smoker) by default, but the engine accepts any
:class:`pricing_projection.MortalityTableQx` /
:class:`pricing_projection.MortalityTableRP2014MP2016` /
:class:`mortality_2017_cso.MortalityTable2017CSO`.

Cashflow shape (Phase 4, Section 3 of ``docs/seven_product_rollout_plan.md``):

* **Death cashflow** at month ``t`` for ``t in [1, horizon_months]``:
  ``face_amount * P(death in month t)``.
* **Single premium** (returned in ``result.single_premium``) is the
  PV of all death-benefit cashflows + PV of monthly expenses (the
  classic NSP_x).

Capabilities: ``supports_economic_scenario=False``,
``supports_monte_carlo=False`` -- WL is fully deterministic.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

import numpy as np

from annuity_model import pricing_projection as sp
from annuity_model._observability import traced
from annuity_model.lapse import LapseAssumption
from annuity_model.mortality_2017_cso import MortalityTable2017CSO


@dataclass(frozen=True)
class WLContract:
    issue_age: int
    sex: Literal["male", "female"]
    smoker_class: Literal["nonsmoker", "smoker"] = "nonsmoker"
    face_amount: float = 250_000.0
    horizon_age: int = 120
    payment_freq_per_year: int = 12

    @property
    def benefit_annual(self) -> float:
        return float(self.face_amount)


@dataclass(frozen=True)
class WLProjectionResult:
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
    expected_premium_cashflows: np.ndarray
    expected_claim_cashflows: np.ndarray
    face_amount: float
    smoker_class: str


def _resolve_wl_mortality(
    mortality: sp.MortalityTableQx | sp.MortalityTableRP2014MP2016 | MortalityTable2017CSO | None,
    *,
    sex: str,
    smoker_class: str,
) -> sp.MortalityTableQx | sp.MortalityTableRP2014MP2016:
    if mortality is None:
        # Default to CSO 2017 placeholder -- the four life products use this
        # by default per the rollout plan.
        cso = MortalityTable2017CSO.load(
            sex=sex,  # type: ignore[arg-type]
            smoker_class=smoker_class,  # type: ignore[arg-type]
        )
        return cso.table
    if isinstance(mortality, MortalityTable2017CSO):
        return mortality.table
    return mortality


@traced("pricing.wl.deterministic")
def price_wl_single_premium(
    *,
    contract: WLContract,
    yield_curve: sp.YieldCurve,
    mortality: sp.MortalityTableQx
    | sp.MortalityTableRP2014MP2016
    | MortalityTable2017CSO
    | None = None,
    horizon_age: int | None = None,
    spread: float = 0.0,
    valuation_year: int | None = None,
    expenses: sp.ExpenseAssumptions | None = None,
    expenses_csv_path: str = sp.DEFAULT_EXPENSES_CSV,
    index_scenario_csv_path: str | None = None,
    expense_annual_inflation: float = 0.0,
    lapse: LapseAssumption | None = None,
) -> WLProjectionResult:
    if contract.payment_freq_per_year != 12:
        raise ValueError("WL scaffold assumes monthly frequency.")
    if contract.face_amount <= 0:
        raise ValueError("face_amount must be > 0.")
    del index_scenario_csv_path  # accepted for adapter-signature parity, unused.

    mort = _resolve_wl_mortality(mortality, sex=contract.sex, smoker_class=contract.smoker_class)

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

    if lapse is not None:
        from annuity_model.lapse import combined_monthly_survival

        mort_step = np.empty(n_months, dtype=float)
        mort_step[0] = survival_end[0]
        mort_step[1:] = survival_end[1:] / np.clip(survival_end[:-1], 1e-15, None)
        mortality_q_m = np.clip(1.0 - mort_step, 0.0, 1.0)
        lapse_q_m = lapse.monthly_decrements(n_months)
        survival_combined = combined_monthly_survival(
            mortality_monthly_q=mortality_q_m,
            lapse_monthly_q=lapse_q_m,
        )
        survival_combined_start = np.empty_like(survival_combined)
        survival_combined_start[0] = 1.0
        survival_combined_start[1:] = survival_combined[:-1]
        death_prob_month = np.clip(survival_combined_start * mortality_q_m, 0.0, 1.0)
        survival_end = survival_combined
        survival_start = survival_combined_start

    death_cf = float(contract.face_amount) * death_prob_month

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
    single_premium = pv_benefit + pv_monthly_expenses

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
    benefit_nominal = np.full(n_months, float(contract.face_amount), dtype=float)
    # Premium cashflow is single premium at issue (positive) — schedule for
    # adapter compatibility shows zero monthly premium (single premium only).
    expected_premium_cashflows = np.zeros(n_months, dtype=float)

    return WLProjectionResult(
        months=months,
        times_years=times_years,
        ages_at_payment=ages_at_payment,
        survival_to_payment=survival_end,
        discount_factors=df,
        pv_benefit=pv_benefit,
        pv_monthly_expenses=pv_monthly_expenses,
        annuity_factor=annuity_factor,
        single_premium=single_premium,
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
        expected_premium_cashflows=expected_premium_cashflows,
        expected_claim_cashflows=expected_claim_cashflows,
        face_amount=float(contract.face_amount),
        smoker_class=str(contract.smoker_class),
    )


def liability_path_from_wl_projection(pricing: WLProjectionResult) -> sp.LiabilityPath:
    return sp.LiabilityPath(
        times_years=np.asarray(pricing.times_years, dtype=float),
        expected_total_cashflows=np.asarray(pricing.expected_total_cashflows, dtype=float),
    )


from annuity_model.liability_dispatch import register_liability_path_converter  # noqa: E402

register_liability_path_converter("WLProjectionResult", liability_path_from_wl_projection)
