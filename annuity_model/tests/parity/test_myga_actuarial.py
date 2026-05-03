"""Actuarial reasonableness tests for MYGA.

Bands live in ``actuarial_benchmarks.py``; tolerances for closed-form
matches live in ``parity_constants.py``. Failures here mean the engine
is producing actuarially nonsense numbers; **do NOT widen the band**.
"""

from __future__ import annotations

import numpy as np
import pytest

import myga_projection as my
import pricing_projection as sp
from actuarial_benchmarks import (
    MYGA_BENCHMARK_AV_T_HI,
    MYGA_BENCHMARK_AV_T_LO,
    MYGA_BENCHMARK_PV_HI,
    MYGA_BENCHMARK_PV_LO,
    MYGA_CLOSED_FORM_AV_TOL,
    MYGA_SENSITIVITY_EPS,
)

pytestmark = [pytest.mark.parity, pytest.mark.product_myga]


def _baseline_contract() -> my.MYGAContract:
    return my.MYGAContract(
        issue_age=60,
        sex="male",
        single_premium=100_000.0,
        declared_rate_annual=0.045,
        guarantee_years=5,
    )


def _baseline_yc(rate: float = 0.045) -> sp.YieldCurve:
    return sp.YieldCurve.from_flat_rate(rate)


def _baseline_mort() -> sp.MortalityTableQx:
    ages = np.arange(0, 121, dtype=int)
    qx = np.clip(0.005 + ages * 1e-5, 1e-6, 0.4)
    return sp.MortalityTableQx(ages, qx)


def _price(contract=None, yc=None, mort=None):
    return my.price_myga_single_premium(
        contract=contract or _baseline_contract(),
        yield_curve=yc or _baseline_yc(),
        mortality=mort or _baseline_mort(),
        horizon_age=70,
        expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
    )


def test_myga_actuarial_sanity_signs():
    """Universal sanity (Section 13.2): all PVs / probs / AV non-negative."""
    res = _price()
    assert res.single_premium >= 0
    assert res.pv_benefit >= 0
    assert (res.survival_to_payment >= 0).all()
    assert (res.survival_to_payment <= 1).all()
    diffs = np.diff(res.survival_to_payment, prepend=1.0)
    assert diffs.max() <= 1e-10  # survival monotone non-increasing
    assert (res.discount_factors > 0).all()
    assert (res.account_value_end_month >= 0).all()


def test_myga_av_within_benchmark_band():
    """Section 13.3: MYGA AV at maturity within $124,500-$124,800."""
    res = _price()
    av_t = float(res.account_value_end_month[-1])
    assert MYGA_BENCHMARK_AV_T_LO <= av_t <= MYGA_BENCHMARK_AV_T_HI, (
        f"MYGA AV(T)={av_t:,.2f} fell outside the band "
        f"[{MYGA_BENCHMARK_AV_T_LO:,.0f}, {MYGA_BENCHMARK_AV_T_HI:,.0f}]. "
        "See docs/actuarial_benchmarks.md row 'MYGA' and Section 13.7 of "
        "the rollout plan for the resolution playbook."
    )


def test_myga_pv_within_benchmark_band():
    """If discount = declared rate, PV(maturity) ≈ premium × survival(T)."""
    res = _price()
    assert MYGA_BENCHMARK_PV_LO <= res.pv_benefit <= MYGA_BENCHMARK_PV_HI, (
        f"MYGA PV(benefit)={res.pv_benefit:,.2f} fell outside the band "
        f"[{MYGA_BENCHMARK_PV_LO:,.0f}, {MYGA_BENCHMARK_PV_HI:,.0f}]."
    )


def test_myga_closed_form_av_match():
    """Section 13.5: AV(T) must equal SP × (1+i)^T to MYGA_CLOSED_FORM_AV_TOL."""
    res = _price()
    av_t = float(res.account_value_end_month[-1])
    closed_form = 100_000.0 * (1.045**5)
    assert abs(av_t - closed_form) <= MYGA_CLOSED_FORM_AV_TOL, (
        f"AV(T)={av_t:.6f} vs closed_form={closed_form:.6f}, "
        f"diff={av_t - closed_form:+.6e} > {MYGA_CLOSED_FORM_AV_TOL}."
    )


def test_myga_yield_sensitivity_negative_pv():
    """Section 13.4: +100bps yield shock must reduce PV(benefit)."""
    base = _price()
    shocked = _price(yc=_baseline_yc(rate=0.045 + 0.01))
    assert shocked.pv_benefit < base.pv_benefit - MYGA_SENSITIVITY_EPS, (
        f"+100bps yield shock did not reduce PV(benefit). "
        f"base={base.pv_benefit:.2f} shocked={shocked.pv_benefit:.2f}. "
        "This is a sign bug."
    )


def test_myga_mortality_reduces_total_cashflow_sum():
    """Heavier mortality should reduce the *undiscounted* sum of cashflows.

    For MYGA, in-period deaths pay AV[t] (less than AV[T]) while maturity
    pays AV[T]. Heavier mortality shifts payments from the higher-AV
    maturity to lower-AV death months, so total CF sum drops. (The
    discounted PV can move either direction depending on how much the
    earlier timing offsets the lower AV — that is engine-correct
    behavior, not a sign bug.)
    """
    base = _price()
    ages = np.arange(0, 121, dtype=int)
    qx_heavy = np.clip(0.05 + ages * 1e-5, 1e-6, 0.4)
    mort_heavy = sp.MortalityTableQx(ages, qx_heavy)
    shocked = _price(mort=mort_heavy)
    base_sum = float(base.expected_total_cashflows.sum())
    shocked_sum = float(shocked.expected_total_cashflows.sum())
    assert shocked_sum < base_sum - MYGA_SENSITIVITY_EPS, (
        f"Heavier mortality should reduce total cashflow sum. "
        f"base={base_sum:.2f} shocked={shocked_sum:.2f}."
    )
