"""Actuarial reasonableness tests for VA."""

from __future__ import annotations

import numpy as np
import pytest

from annuity_model import pricing_projection as sp
from annuity_model import va_projection as va
from annuity_model.actuarial_benchmarks import (
    VA_BENCHMARK_AV_T_FLAT_HI,
    VA_BENCHMARK_AV_T_FLAT_LO,
    VA_BENCHMARK_AV_T_MC_HI,
    VA_BENCHMARK_AV_T_MC_LO,
    VA_SENSITIVITY_EPS,
)
from annuity_model.parity_constants import AV_TOL

pytestmark = [pytest.mark.parity, pytest.mark.product_va]


def _baseline_contract() -> va.VAContract:
    return va.VAContract(
        issue_age=55,
        sex="male",
        single_premium=100_000.0,
        me_charge_annual=0.014,
        horizon_years=20,
    )


def _yc(r: float = 0.04) -> sp.YieldCurve:
    return sp.YieldCurve.from_flat_rate(r)


def _mort() -> sp.MortalityTableQx:
    ages = np.arange(0, 121, dtype=int)
    qx = np.clip(0.005 + ages * 1e-5, 1e-6, 0.4)
    return sp.MortalityTableQx(ages, qx)


def test_va_actuarial_sanity_signs():
    contract = _baseline_contract()
    n = int(contract.horizon_years * 12)
    rng = np.random.default_rng(42)
    levels = 100.0 * np.cumprod(1.0 + rng.normal(0.005, 0.03, size=n))
    res = va.price_va_single_premium(
        contract=contract,
        yield_curve=_yc(),
        mortality=_mort(),
        horizon_age=75,
        expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
        index_s0=100.0,
        index_levels_payment=levels,
    )
    assert res.single_premium >= 0
    assert res.pv_benefit >= 0
    assert (res.account_value_end_month >= 0).all()
    assert (res.discount_factors > 0).all()


def test_va_av_flat_path_within_band():
    """Flat S&P (no growth): 20y of 1.4% M&E shrinks AV to ~75k."""
    contract = _baseline_contract()
    n = int(contract.horizon_years * 12)
    levels = np.full(n, 100.0)  # flat
    res = va.price_va_single_premium(
        contract=contract,
        yield_curve=_yc(),
        mortality=_mort(),
        horizon_age=75,
        expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
        index_s0=100.0,
        index_levels_payment=levels,
    )
    av_t = float(res.account_value_end_month[-1])
    assert VA_BENCHMARK_AV_T_FLAT_LO <= av_t <= VA_BENCHMARK_AV_T_FLAT_HI, (
        f"VA AV(T) under flat-path={av_t:,.2f} fell outside band "
        f"[{VA_BENCHMARK_AV_T_FLAT_LO:,.0f}, {VA_BENCHMARK_AV_T_FLAT_HI:,.0f}]."
    )


def test_va_collapses_when_me_zero_and_flat():
    """Section 13.5: M&E=0, flat S&P -> AV(T) ≈ premium."""
    contract = va.VAContract(
        issue_age=55,
        sex="male",
        single_premium=100_000.0,
        me_charge_annual=0.0,
        horizon_years=10,
    )
    n = 120
    levels = np.full(n, 100.0)
    res = va.price_va_single_premium(
        contract=contract,
        yield_curve=_yc(),
        mortality=_mort(),
        horizon_age=65,
        expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
        index_s0=100.0,
        index_levels_payment=levels,
    )
    av_t = float(res.account_value_end_month[-1])
    assert abs(av_t - 100_000.0) <= AV_TOL, (
        f"VA with M&E=0 and flat S&P should preserve premium; got AV(T)={av_t:.6f}."
    )


def test_va_mc_mean_within_band():
    """Section 13.3 MC band $170k-$320k for 20y, 6% drift, 1.4% M&E."""
    contract = _baseline_contract()
    res = va.price_va_single_premium_monte_carlo(
        contract=contract,
        yield_curve=_yc(),
        mortality=_mort(),
        horizon_age=75,
        expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
        n_sims=80,
        annual_drift=0.06,
        annual_vol=0.15,
        seed=42,
        s0=100.0,
    )
    assert VA_BENCHMARK_AV_T_MC_LO <= res.av_end_mean <= VA_BENCHMARK_AV_T_MC_HI, (
        f"VA E[AV(T)] MC mean={res.av_end_mean:,.2f} fell outside band "
        f"[{VA_BENCHMARK_AV_T_MC_LO:,.0f}, {VA_BENCHMARK_AV_T_MC_HI:,.0f}]."
    )


def test_va_drift_increase_raises_av():
    """Section 13.4: higher sub-account drift -> higher expected AV."""
    contract = _baseline_contract()
    base = va.price_va_single_premium_monte_carlo(
        contract=contract,
        yield_curve=_yc(),
        mortality=_mort(),
        horizon_age=75,
        expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
        n_sims=40,
        annual_drift=0.06,
        annual_vol=0.15,
        seed=42,
        s0=100.0,
    )
    higher = va.price_va_single_premium_monte_carlo(
        contract=contract,
        yield_curve=_yc(),
        mortality=_mort(),
        horizon_age=75,
        expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
        n_sims=40,
        annual_drift=0.10,
        annual_vol=0.15,
        seed=42,
        s0=100.0,
    )
    assert higher.av_end_mean > base.av_end_mean + VA_SENSITIVITY_EPS, (
        f"Higher drift should raise E[AV]. base={base.av_end_mean:.2f} "
        f"higher={higher.av_end_mean:.2f}."
    )
