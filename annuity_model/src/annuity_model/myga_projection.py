"""MYGA (Multi-Year Guaranteed Annuity) pricing engine.

A single-premium fixed deferred annuity. The issuer guarantees a
declared annual rate for ``guarantee_years``; at end of the guarantee
period the contract matures (no surrender / re-rating modeling in v1).

Cashflow shape (Phase 1, Section 3 of ``docs/seven_product_rollout_plan.md``):

* **Maturity payout** at month ``T = guarantee_years * 12``:
  ``AV[T] * survival[T-1]`` (alive-weighted).
* **In-period death payout** at month ``t`` for ``t in [1, T)``:
  ``AV[t] * P(death in month t)`` where
  ``P(death in month t) = survival[t-1] - survival[t]``.
* **Optional lapse payout** when ``lapse=`` is supplied:
  ``AV[t] * P(lapse in month t)`` (no surrender charge in v1's
  pricing even if a schedule is recorded for display).

The engine returns the pre-computed accumulation path plus the
cashflow vector; PV / pricing / liability path follow the same shape
as the other accumulation products (RILA, FIA, VA).

Capabilities: ``supports_economic_scenario=False``,
``supports_monte_carlo=False`` -- MYGA pricing is fully deterministic
given the contract.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

import numpy as np

from annuity_model import pricing_projection as sp
from annuity_model._observability import traced
from annuity_model.lapse import LapseAssumption


@dataclass(frozen=True)
class MYGAContract:
    """Single-premium MYGA contract (v1: no riders, no surrender modeling in pricing)."""

    issue_age: int
    sex: Literal["male", "female"]
    single_premium: float
    declared_rate_annual: float
    guarantee_years: int = 5
    payment_freq_per_year: int = 12

    @property
    def benefit_annual(self) -> float:
        """UI / export compatibility (no SPIA-style annual benefit)."""
        return 0.0


@dataclass(frozen=True)
class MYGAProjectionResult:
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
    expected_claim_cashflows: np.ndarray
    declared_rate_annual: float
    guarantee_years: int


@traced("pricing.myga.deterministic")
def price_myga_single_premium(
    *,
    contract: MYGAContract,
    yield_curve: sp.YieldCurve,
    mortality: sp.MortalityTableQx | sp.MortalityTableRP2014MP2016,
    horizon_age: int,
    spread: float = 0.0,
    valuation_year: int | None = None,
    expenses: sp.ExpenseAssumptions | None = None,
    expenses_csv_path: str = sp.DEFAULT_EXPENSES_CSV,
    index_scenario_csv_path: str | None = None,
    expense_annual_inflation: float = 0.0,
    lapse: LapseAssumption | None = None,
) -> MYGAProjectionResult:
    """Price a single-premium MYGA, returning the full projection bundle.

    ``index_scenario_csv_path`` is accepted for adapter signature parity
    but ignored (MYGA is not index-linked). ``expenses`` are used only
    for the PV(expenses) attribute; pricing does not solve for premium
    -- the input ``contract.single_premium`` is the single premium.
    """
    if contract.payment_freq_per_year != 12:
        raise ValueError("MYGA scaffold assumes monthly frequency.")
    if contract.guarantee_years <= 0:
        raise ValueError("guarantee_years must be positive.")
    if contract.single_premium <= 0:
        raise ValueError("single_premium must be > 0.")
    if contract.declared_rate_annual <= -1.0:
        raise ValueError("declared_rate_annual must be > -1.")
    del index_scenario_csv_path  # accepted for adapter-signature parity, unused.

    dt = 1.0 / 12.0
    max_model_months = max(1, int(round((horizon_age - contract.issue_age) / dt)))
    term_months = int(contract.guarantee_years * 12)
    n_months = max(1, min(max_model_months, term_months))

    months = np.arange(1, n_months + 1, dtype=int)
    times_years = months * dt
    ages_at_payment = contract.issue_age + times_years

    if valuation_year is None and isinstance(mortality, sp.MortalityTableRP2014MP2016):
        raise ValueError("valuation_year must be provided when using MortalityTableRP2014MP2016.")
    survival_end = mortality.monthly_survival_to_payment(
        issue_age=contract.issue_age,
        n_months=n_months,
        valuation_year=valuation_year,
    )
    survival_start = np.empty_like(survival_end)
    survival_start[0] = 1.0
    survival_start[1:] = survival_end[:-1]
    death_prob_month = np.clip(survival_start - survival_end, 0.0, 1.0)

    # AV grows monthly at (1 + i)^(t/12) -- continuous-equivalent monthly
    # accumulation; AV[t] is end-of-month-t value.
    monthly_growth = (1.0 + float(contract.declared_rate_annual)) ** dt
    av_end = float(contract.single_premium) * monthly_growth ** np.arange(1, n_months + 1)

    # Optional lapse decrement
    lapse_prob_month = np.zeros(n_months, dtype=float)
    if lapse is not None:
        lapse_q_m = lapse.monthly_decrements(n_months)
        # Combined survival under independent decrements.
        from annuity_model.lapse import combined_monthly_survival, monthly_mortality_q_from_annual

        # Convert mortality monthly survivals to monthly q.
        mort_survival_step = np.empty(n_months, dtype=float)
        mort_survival_step[0] = survival_end[0]
        mort_survival_step[1:] = survival_end[1:] / np.clip(survival_end[:-1], 1e-15, None)
        mortality_q_m = np.clip(1.0 - mort_survival_step, 0.0, 1.0)
        del monthly_mortality_q_from_annual  # imported for re-export, unused locally
        survival_combined = combined_monthly_survival(
            mortality_monthly_q=mortality_q_m,
            lapse_monthly_q=lapse_q_m,
        )
        survival_combined_start = np.empty_like(survival_combined)
        survival_combined_start[0] = 1.0
        survival_combined_start[1:] = survival_combined[:-1]
        # Death prob and lapse prob attributable per month.
        death_prob_month = np.clip(survival_combined_start * mortality_q_m, 0.0, 1.0)
        lapse_prob_month = np.clip(
            survival_combined_start * (1.0 - mortality_q_m) * lapse_q_m, 0.0, 1.0
        )
        # Override survival_end to use the combined (mortality+lapse) decrement.
        survival_end = survival_combined
        survival_start = survival_combined_start

    # In-period death cashflow (pays AV at t for fraction P(death this month)).
    death_cf = av_end * death_prob_month
    # Optional lapse cashflow (no surrender charge in v1).
    lapse_cf = av_end * lapse_prob_month
    # Maturity cashflow at month T (pays AV[T] alive-weighted).
    maturity_cf = np.zeros(n_months, dtype=float)
    maturity_cf[-1] = float(av_end[-1] * survival_end[-1])

    expected_benefit_cashflows = death_cf + lapse_cf + maturity_cf
    expected_claim_cashflows = death_cf  # pure-mortality slice for reporting

    # Expenses (alive-weighted; same convention as existing engines).
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
    expected_total_cashflows = expected_benefit_cashflows + expected_expense_cashflows

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
    benefit_nominal = np.zeros(n_months, dtype=float)
    benefit_nominal[-1] = float(av_end[-1])

    return MYGAProjectionResult(
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
        account_value_end_month=av_end,
        expected_claim_cashflows=expected_claim_cashflows,
        declared_rate_annual=float(contract.declared_rate_annual),
        guarantee_years=int(contract.guarantee_years),
    )


def liability_path_from_myga_projection(pricing: MYGAProjectionResult) -> sp.LiabilityPath:
    return sp.LiabilityPath(
        times_years=np.asarray(pricing.times_years, dtype=float),
        expected_total_cashflows=np.asarray(pricing.expected_total_cashflows, dtype=float),
    )


# Register with the liability-path dispatch so the engine core can route
# `run_alm_projection_from_pricing_result` without an isinstance chain.
from annuity_model.liability_dispatch import register_liability_path_converter  # noqa: E402

register_liability_path_converter("MYGAProjectionResult", liability_path_from_myga_projection)
