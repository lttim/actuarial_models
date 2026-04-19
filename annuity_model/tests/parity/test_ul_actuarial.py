"""Actuarial reasonableness tests for Universal Life."""

from __future__ import annotations

import numpy as np
import pytest

import pricing_projection as sp
import ul_projection as ul
from actuarial_benchmarks import (
    UL_BENCHMARK_AV_20Y_HI,
    UL_BENCHMARK_AV_20Y_LO,
    UL_SENSITIVITY_EPS,
)
from parity_constants import AV_TOL

pytestmark = [pytest.mark.parity, pytest.mark.product_ul]


def _baseline_contract() -> ul.ULContract:
    return ul.ULContract(
        issue_age=45, sex="male", face_amount=250_000.0, single_premium=25_000.0,
        premium_load_pct=0.06, monthly_expense_charge=7.50, declared_rate_annual=0.04,
    )


def _yc(r: float = 0.04) -> sp.YieldCurve:
    return sp.YieldCurve.from_flat_rate(r)


def _mort() -> sp.MortalityTableQx:
    ages = np.arange(0, 121, dtype=int)
    qx = np.clip(0.005 + ages * 1e-5, 1e-6, 0.4)
    return sp.MortalityTableQx(ages, qx)


def _price(contract=None, yc=None, mort=None):
    return ul.price_ul_single_premium(
        contract=contract or _baseline_contract(),
        yield_curve=yc or _yc(), mortality=mort or _mort(),
        horizon_age=120, expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
    )


def test_ul_actuarial_sanity_signs():
    res = _price()
    assert res.single_premium > 0
    assert res.pv_benefit >= 0
    assert (res.account_value_end_month >= 0).all()
    # COI <= qx*face per month (NAR <= face).
    assert (res.coi_dollars >= 0).all()
    assert (res.discount_factors > 0).all()


def test_ul_av_at_20y_within_band():
    res = _price()
    if res.account_value_end_month.size >= 240:
        av_20y = float(res.account_value_end_month[239])
    else:
        av_20y = float(res.account_value_end_month[-1])
    assert UL_BENCHMARK_AV_20Y_LO <= av_20y <= UL_BENCHMARK_AV_20Y_HI, (
        f"UL AV(20y)={av_20y:,.2f} fell outside band "
        f"[{UL_BENCHMARK_AV_20Y_LO:,.0f}, {UL_BENCHMARK_AV_20Y_HI:,.0f}]."
    )


def test_ul_nar_zero_when_av_geq_face():
    """If AV >= face (Type A), NAR == 0 and COI == 0."""
    big_premium = ul.ULContract(
        issue_age=45, sex="male", face_amount=10_000.0, single_premium=100_000.0,
        premium_load_pct=0.0, monthly_expense_charge=0.0, declared_rate_annual=0.04,
    )
    res = _price(contract=big_premium)
    # AV[0] = 100_000; face = 10_000 -> NAR = 0 throughout.
    np.testing.assert_allclose(res.nar_end_month, 0.0, atol=1e-9)
    np.testing.assert_allclose(res.coi_dollars, 0.0, atol=1e-9)


def test_ul_higher_premium_extends_av():
    """Larger SP should make AV last longer (more months before depletion)."""
    base = _price()
    bigger = ul.ULContract(
        issue_age=45, sex="male", face_amount=250_000.0, single_premium=50_000.0,
        premium_load_pct=0.06, monthly_expense_charge=7.50, declared_rate_annual=0.04,
    )
    res_big = _price(contract=bigger)
    base_terminate = (
        int(np.argmax(base.is_terminated_after_month))
        if any(base.is_terminated_after_month)
        else len(base.is_terminated_after_month)
    )
    big_terminate = (
        int(np.argmax(res_big.is_terminated_after_month))
        if any(res_big.is_terminated_after_month)
        else len(res_big.is_terminated_after_month)
    )
    assert big_terminate >= base_terminate, (
        f"Larger premium should extend AV survival. "
        f"base terminates at month {base_terminate}, "
        f"bigger terminates at month {big_terminate}."
    )


def test_ul_yield_sensitivity_negative_pv():
    base = _price()
    shocked = _price(yc=_yc(0.05))
    # PV claims should drop at higher discount rate.
    assert shocked.pv_benefit < base.pv_benefit - UL_SENSITIVITY_EPS, (
        f"+100bps shock did not reduce UL PV. "
        f"base={base.pv_benefit:.2f} shocked={shocked.pv_benefit:.2f}."
    )
