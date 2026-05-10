"""VA (Variable Annuity) pricing engine.

Single-premium deferred VA with a sub-account modeled as the index path
(deterministic CSV by default; GBM-simulated under Monte Carlo). M&E
charge taken monthly from AV. GMDB = max(AV, return_of_premium) at death.

Cashflow shape (Phase 3, Section 3 of ``docs/seven_product_rollout_plan.md``):

* **Maturity payout** at month ``T = horizon_years * 12``:
  ``AV[T] * survival[T-1]`` (alive-weighted).
* **In-period death payout (GMDB)** at month ``t``:
  ``max(AV[t], single_premium) * P(death in month t)``.
* **Optional lapse payout** when ``lapse=`` is supplied:
  ``AV[t] * P(lapse in month t)`` (no surrender charge in v1).

Capabilities: ``supports_economic_scenario=True``,
``supports_monte_carlo=True``.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

import numpy as np

from annuity_model import pricing_projection as sp
from annuity_model._observability import traced
from annuity_model.lapse import LapseAssumption


@dataclass(frozen=True)
class VAContract:
    issue_age: int
    sex: Literal["male", "female"]
    single_premium: float
    me_charge_annual: float = 0.014
    gmdb_basis: Literal["return_of_premium", "max_anniversary"] = "return_of_premium"
    horizon_years: int = 20
    payment_freq_per_year: int = 12
    subaccount_drift_annual: float = 0.06
    subaccount_vol_annual: float = 0.15

    @property
    def benefit_annual(self) -> float:
        return 0.0


@dataclass(frozen=True)
class VAProjectionResult:
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
    horizon_years: int


@traced("pricing.va.deterministic")
def price_va_single_premium(
    *,
    contract: VAContract,
    yield_curve: sp.YieldCurve,
    mortality: sp.MortalityTableQx | sp.MortalityTableRP2014MP2016,
    horizon_age: int,
    spread: float = 0.0,
    valuation_year: int | None = None,
    expenses: sp.ExpenseAssumptions | None = None,
    expenses_csv_path: str = sp.DEFAULT_EXPENSES_CSV,
    index_scenario_csv_path: str | None = None,
    index_s0: float | None = None,
    index_levels_payment: np.ndarray | None = None,
    expense_annual_inflation: float = 0.0,
    lapse: LapseAssumption | None = None,
) -> VAProjectionResult:
    if contract.payment_freq_per_year != 12:
        raise ValueError("VA scaffold assumes monthly frequency.")
    if contract.single_premium <= 0:
        raise ValueError("single_premium must be > 0.")
    if not (0.0 <= contract.me_charge_annual < 1.0):
        raise ValueError("me_charge_annual must be in [0, 1).")

    dt = 1.0 / 12.0
    n_months = max(
        1,
        min(
            int(round((horizon_age - contract.issue_age) / dt)),
            int(contract.horizon_years * 12),
        ),
    )

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

    me_monthly = float(contract.me_charge_annual) / 12.0
    av = float(contract.single_premium)
    av_end = np.zeros(n_months, dtype=float)
    L = np.zeros(n_months + 1, dtype=float)
    L[0] = float(s0)
    L[1:] = np.asarray(levels_payment, dtype=float)
    for t in range(1, n_months + 1):
        ratio = float(L[t] / L[t - 1]) if L[t - 1] > 0 else 1.0
        av = av * ratio * (1.0 - me_monthly)
        av_end[t - 1] = max(0.0, av)

    lapse_prob_month = np.zeros(n_months, dtype=float)
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
        lapse_prob_month = np.clip(
            survival_combined_start * (1.0 - mortality_q_m) * lapse_q_m, 0.0, 1.0
        )
        survival_end = survival_combined
        survival_start = survival_combined_start

    # GMDB death-benefit cashflow
    if contract.gmdb_basis == "return_of_premium":
        db = np.maximum(av_end, float(contract.single_premium))
    else:
        # max_anniversary: high-water mark
        db = np.maximum.accumulate(av_end)
        db = np.maximum(db, float(contract.single_premium))

    death_cf = db * death_prob_month
    lapse_cf = av_end * lapse_prob_month
    maturity_cf = np.zeros(n_months, dtype=float)
    maturity_cf[-1] = float(av_end[-1] * survival_end[-1])

    expected_benefit_cashflows = death_cf + lapse_cf + maturity_cf
    expected_claim_cashflows = death_cf

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

    return VAProjectionResult(
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
        benefit_nominal_scheduled=expected_benefit_cashflows,
        expense_nominal_scheduled=expected_expense_cashflows,
        expense_annual_inflation=float(expense_annual_inflation),
        index_s0=float(s0),
        account_value_end_month=av_end,
        expected_claim_cashflows=expected_claim_cashflows,
        horizon_years=int(contract.horizon_years),
    )


@dataclass(frozen=True)
class VAMonteCarloResult:
    n_sims: int
    pv_benefit: np.ndarray
    pv_total_mean: float
    pv_benefit_mean: float
    av_end_mean: float


@traced("pricing.va.monte_carlo")
def price_va_single_premium_monte_carlo(
    *,
    contract: VAContract,
    yield_curve: sp.YieldCurve,
    mortality: sp.MortalityTableQx | sp.MortalityTableRP2014MP2016,
    horizon_age: int,
    spread: float = 0.0,
    valuation_year: int | None = None,
    expenses: sp.ExpenseAssumptions | None = None,
    expenses_csv_path: str = sp.DEFAULT_EXPENSES_CSV,
    expense_annual_inflation: float = 0.0,
    n_sims: int = 500,
    annual_drift: float = 0.06,
    annual_vol: float = 0.15,
    seed: int | None = None,
    s0: float = 100.0,
) -> VAMonteCarloResult:
    dt = 1.0 / 12.0
    n_months = max(
        1,
        min(
            int(round((horizon_age - contract.issue_age) / dt)),
            int(contract.horizon_years * 12),
        ),
    )
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
        res = price_va_single_premium(
            contract=contract,
            yield_curve=yield_curve,
            mortality=mortality,
            horizon_age=horizon_age,
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
    return VAMonteCarloResult(
        n_sims=int(n_sims),
        pv_benefit=pvb,
        pv_total_mean=float(np.nanmean(pvb)),
        pv_benefit_mean=float(np.nanmean(pvb)),
        av_end_mean=float(np.nanmean(av_end)),
    )


def liability_path_from_va_projection(pricing: VAProjectionResult) -> sp.LiabilityPath:
    return sp.LiabilityPath(
        times_years=np.asarray(pricing.times_years, dtype=float),
        expected_total_cashflows=np.asarray(pricing.expected_total_cashflows, dtype=float),
    )


from annuity_model.liability_dispatch import register_liability_path_converter  # noqa: E402

register_liability_path_converter("VAProjectionResult", liability_path_from_va_projection)
