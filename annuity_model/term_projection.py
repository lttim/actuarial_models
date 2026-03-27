from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

import numpy as np

import pricing_projection as sp


@dataclass(frozen=True)
class TermLifeContract:
    issue_age: int
    sex: Literal["male", "female"]
    death_benefit: float
    monthly_premium: float = 250.0
    term_years: int = 20
    premium_mode: Literal["level_monthly"] = "level_monthly"
    benefit_timing: Literal["eoy_death"] = "eoy_death"
    payment_freq_per_year: int = 12

    @property
    def benefit_annual(self) -> float:
        # UI/reporting compatibility with existing SPIA result renderers.
        return float(self.death_benefit)


@dataclass(frozen=True)
class TermLifeProjectionResult:
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


def price_term_life_level_monthly(
    *,
    contract: TermLifeContract,
    yield_curve: sp.YieldCurve,
    mortality: sp.MortalityTableQx | sp.MortalityTableRP2014MP2016,
    horizon_age: int,
    spread: float = 0.0,
    valuation_year: int | None = None,
) -> TermLifeProjectionResult:
    if contract.payment_freq_per_year != 12:
        raise ValueError("Term scaffold currently assumes monthly frequency.")
    if contract.term_years <= 0:
        raise ValueError("term_years must be positive.")
    if contract.monthly_premium < 0.0:
        raise ValueError("monthly_premium must be non-negative.")
    if contract.death_benefit <= 0.0:
        raise ValueError("death_benefit must be positive.")

    dt = 1.0 / 12.0
    max_model_months = max(1, int(round((horizon_age - contract.issue_age) / dt)))
    term_months = int(contract.term_years * 12)
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

    claim_cf = np.zeros(n_months, dtype=float)
    if contract.benefit_timing != "eoy_death":
        raise ValueError("Only eoy_death timing is supported in this release.")
    for year_end_m in range(12, n_months + 1, 12):
        start = year_end_m - 12
        end = year_end_m
        claim_cf[year_end_m - 1] = float(contract.death_benefit * np.sum(death_prob_month[start:end]))

    premium_cf = np.full(n_months, float(contract.monthly_premium), dtype=float) * survival_start
    expected_total = claim_cf - premium_cf

    df = yield_curve.discount_factors(times_years, spread=spread)
    pv_claims = float(np.sum(claim_cf * df))
    pv_premiums = float(np.sum(premium_cf * df))
    net_pv = pv_claims - pv_premiums

    reserve_times_years = np.concatenate(([0.0], times_years))
    economic_reserve = np.zeros(n_months + 1, dtype=float)
    pv_remaining = np.zeros(n_months + 1, dtype=float)
    for i in range(n_months - 1, -1, -1):
        pv_remaining[i] = float(expected_total[i] * df[i] + pv_remaining[i + 1])
    economic_reserve[0] = float(pv_remaining[0])
    for i in range(1, n_months + 1):
        denom = max(df[i - 1], 1e-15)
        economic_reserve[i] = float(pv_remaining[i] / denom) if i < n_months else 0.0

    idx_level = np.full(n_months, 100.0, dtype=float)
    idx_zero = np.zeros(n_months, dtype=float)
    benefit_nominal = np.where((months % 12) == 0, float(contract.death_benefit), 0.0)

    annuity_factor = float(np.sum(survival_start * df))
    return TermLifeProjectionResult(
        months=months,
        times_years=times_years,
        ages_at_payment=ages_at_payment,
        survival_to_payment=survival_end,
        discount_factors=df,
        pv_benefit=pv_claims,
        pv_monthly_expenses=-pv_premiums,
        annuity_factor=annuity_factor,
        single_premium=net_pv,
        expected_benefit_cashflows=claim_cf,
        expected_expense_cashflows=-premium_cf,
        expected_total_cashflows=expected_total,
        reserve_times_years=reserve_times_years,
        economic_reserve=economic_reserve,
        index_level_at_payment=idx_level,
        index_simple_return=idx_zero.copy(),
        index_log_return=idx_zero.copy(),
        index_cumulative_return=idx_zero.copy(),
        benefit_nominal_scheduled=benefit_nominal,
        expense_nominal_scheduled=-np.full(n_months, float(contract.monthly_premium), dtype=float),
        expense_annual_inflation=0.0,
        index_s0=100.0,
        expected_premium_cashflows=premium_cf,
        expected_claim_cashflows=claim_cf,
    )


def liability_path_from_term_projection(pricing: TermLifeProjectionResult) -> sp.LiabilityPath:
    return sp.LiabilityPath(
        times_years=np.asarray(pricing.times_years, dtype=float),
        expected_total_cashflows=np.asarray(pricing.expected_total_cashflows, dtype=float),
    )
