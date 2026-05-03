"""Actuarial reasonableness tests for FIA."""

from __future__ import annotations

import numpy as np
import pytest

import fia_projection as fp
import pricing_projection as sp
from actuarial_benchmarks import (
    FIA_BENCHMARK_AV_T_HI,
    FIA_BENCHMARK_AV_T_LO,
    FIA_SENSITIVITY_EPS,
)
from parity_constants import AV_TOL

pytestmark = [pytest.mark.parity, pytest.mark.product_fia]


def _baseline_contract() -> fp.FIAContract:
    return fp.FIAContract(
        issue_age=60,
        sex="male",
        single_premium=100_000.0,
        participation=0.8,
        cap=0.07,
        floor=0.0,
        horizon_years=10,
    )


def _yc(r: float = 0.04) -> sp.YieldCurve:
    return sp.YieldCurve.from_flat_rate(r)


def _mort() -> sp.MortalityTableQx:
    ages = np.arange(0, 121, dtype=int)
    qx = np.clip(0.005 + ages * 1e-5, 1e-6, 0.4)
    return sp.MortalityTableQx(ages, qx)


def _price(contract=None, levels=None, yc=None):
    if contract is None:
        contract = _baseline_contract()
    n_months = int(contract.horizon_years * 12)
    if levels is None:
        rng = np.random.default_rng(42)
        levels = 100.0 * np.cumprod(1.0 + rng.normal(0.005, 0.02, size=n_months))
    return fp.price_fia_single_premium(
        contract=contract,
        yield_curve=yc or _yc(),
        mortality=_mort(),
        horizon_age=70,
        expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
        index_s0=100.0,
        index_levels_payment=levels,
    )


def test_fia_actuarial_sanity_signs():
    res = _price()
    assert res.single_premium >= 0
    assert res.pv_benefit >= 0
    assert (res.survival_to_payment >= 0).all()
    assert (res.survival_to_payment <= 1).all()
    assert (res.discount_factors > 0).all()
    assert (res.account_value_end_month >= 0).all()


def test_fia_av_floor_zero_means_av_non_decreasing_at_anniversaries():
    """With floor=0 and rider_fee=0, AV at end of segment >= AV at start."""
    contract = fp.FIAContract(
        issue_age=60,
        sex="male",
        single_premium=100_000.0,
        participation=0.8,
        cap=0.07,
        floor=0.0,
        horizon_years=10,
        rider_fee_annual=0.0,
    )
    res = _price(contract=contract)
    av = res.account_value_end_month
    # End of each segment year (months 12, 24, ..., 120) >= prior segment's end.
    for y in range(2, 11):
        assert av[y * 12 - 1] >= av[(y - 1) * 12 - 1] - 1e-9, (
            f"FIA AV at end of year {y} ({av[y * 12 - 1]:.2f}) is less than year "
            f"{y - 1} ({av[(y - 1) * 12 - 1]:.2f}). Floor=0 should prevent decrease."
        )


def test_fia_av_at_horizon_within_band():
    res = _price()
    av_t = float(res.account_value_end_month[-1])
    assert FIA_BENCHMARK_AV_T_LO <= av_t <= FIA_BENCHMARK_AV_T_HI, (
        f"FIA AV(T)={av_t:,.2f} fell outside band "
        f"[{FIA_BENCHMARK_AV_T_LO:,.0f}, {FIA_BENCHMARK_AV_T_HI:,.0f}]."
    )


def test_fia_collapses_when_cap_equals_floor_zero():
    """Section 13.5: floor = cap = 0 means no growth -> AV stays at premium."""
    contract = fp.FIAContract(
        issue_age=60,
        sex="male",
        single_premium=100_000.0,
        participation=0.8,
        cap=0.0,
        floor=0.0,
        horizon_years=10,
        rider_fee_annual=0.0,
    )
    res = _price(contract=contract)
    av_t = float(res.account_value_end_month[-1])
    assert abs(av_t - 100_000.0) <= AV_TOL, (
        f"FIA with cap=floor=0 should preserve premium; got AV(T)={av_t:.6f}."
    )


def test_fia_cap_increase_increases_av():
    """Section 13.4: higher cap should increase expected AV (with non-negative index)."""
    base_contract = _baseline_contract()
    high_cap_contract = fp.FIAContract(
        issue_age=60,
        sex="male",
        single_premium=100_000.0,
        participation=0.8,
        cap=0.12,
        floor=0.0,
        horizon_years=10,
    )
    rng = np.random.default_rng(42)
    n = 120
    levels = 100.0 * np.cumprod(1.0 + rng.normal(0.01, 0.005, size=n))  # mostly positive
    base = _price(contract=base_contract, levels=levels)
    shocked = _price(contract=high_cap_contract, levels=levels)
    assert (
        shocked.account_value_end_month[-1] > base.account_value_end_month[-1] + FIA_SENSITIVITY_EPS
    ), "Higher cap should increase AV when index has positive cumulative return."
