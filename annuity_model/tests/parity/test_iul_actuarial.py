"""Actuarial reasonableness tests for Indexed UL."""

from __future__ import annotations

import numpy as np
import pytest

import iul_projection as iul
import pricing_projection as sp
import ul_projection as ul
from actuarial_benchmarks import (
    IUL_BENCHMARK_AV_20Y_HI,
    IUL_BENCHMARK_AV_20Y_LO,
    IUL_SENSITIVITY_EPS,
)
from parity_constants import AV_TOL

pytestmark = [pytest.mark.parity, pytest.mark.product_iul]


def _baseline_contract() -> iul.IULContract:
    return iul.IULContract(
        issue_age=45, sex="male", face_amount=250_000.0, single_premium=25_000.0,
        premium_load_pct=0.06, monthly_expense_charge=7.50,
        participation=1.0, cap=0.10, floor=0.0,
    )


def _yc(r: float = 0.04) -> sp.YieldCurve:
    return sp.YieldCurve.from_flat_rate(r)


def _mort() -> sp.MortalityTableQx:
    ages = np.arange(0, 121, dtype=int)
    qx = np.clip(0.005 + ages * 1e-5, 1e-6, 0.4)
    return sp.MortalityTableQx(ages, qx)


def _price(contract=None, levels=None, horizon_age=120):
    contract = contract or _baseline_contract()
    n_months = (horizon_age - contract.issue_age) * 12
    if levels is None:
        rng = np.random.default_rng(42)
        levels = 100.0 * np.cumprod(1.0 + rng.normal(0.005, 0.03, size=n_months))
    return iul.price_iul_single_premium(
        contract=contract, yield_curve=_yc(), mortality=_mort(),
        horizon_age=horizon_age, expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
        index_s0=100.0, index_levels_payment=levels,
    )


def test_iul_actuarial_sanity_signs():
    res = _price()
    assert res.single_premium > 0
    assert res.pv_benefit >= 0
    assert (res.account_value_end_month >= 0).all()
    # Per-segment credits within [floor, cap]
    cred = res.segment_credited_rate
    nonzero = cred[cred != 0]
    assert (nonzero >= -1e-12).all() and (nonzero <= 0.10 + 1e-12).all(), (
        "IUL segment credits must be within [floor, cap]."
    )


def test_iul_av_at_20y_within_band():
    res = _price()
    if res.account_value_end_month.size >= 240:
        av_20y = float(res.account_value_end_month[239])
    else:
        av_20y = float(res.account_value_end_month[-1])
    assert IUL_BENCHMARK_AV_20Y_LO <= av_20y <= IUL_BENCHMARK_AV_20Y_HI, (
        f"IUL AV(20y)={av_20y:,.2f} fell outside band "
        f"[{IUL_BENCHMARK_AV_20Y_LO:,.0f}, {IUL_BENCHMARK_AV_20Y_HI:,.0f}]."
    )


def test_iul_floor_zero_protects_from_index_drops():
    """With floor=0, AV must not decrease at segment anniversaries from index loss."""
    res = _price()
    cred = res.segment_credited_rate
    # No segment credit < 0 (floor=0).
    assert (cred >= -1e-12).all(), (
        f"IUL segment credits below 0 with floor=0; got min={cred.min()}."
    )


def test_iul_cap_zero_floor_zero_collapses_to_no_credit_ul():
    """Section 13.5: cap=floor=0 should match no-credit UL with same params."""
    iul_contract = iul.IULContract(
        issue_age=45, sex="male", face_amount=250_000.0, single_premium=25_000.0,
        premium_load_pct=0.06, monthly_expense_charge=7.50,
        participation=1.0, cap=0.0, floor=0.0,
    )
    ul_contract = ul.ULContract(
        issue_age=45, sex="male", face_amount=250_000.0, single_premium=25_000.0,
        premium_load_pct=0.06, monthly_expense_charge=7.50,
        declared_rate_annual=0.0,
    )
    res_iul = iul.price_iul_single_premium(
        contract=iul_contract, yield_curve=_yc(), mortality=_mort(),
        horizon_age=80, expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
        index_s0=100.0,
        index_levels_payment=np.full((80 - 45) * 12, 100.0),
    )
    res_ul = ul.price_ul_single_premium(
        contract=ul_contract, yield_curve=_yc(), mortality=_mort(),
        horizon_age=80, expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
    )
    # AV paths must match within AV_TOL.
    np.testing.assert_allclose(
        res_iul.account_value_end_month,
        res_ul.account_value_end_month,
        rtol=0.0, atol=AV_TOL,
        err_msg="IUL with cap=floor=0 should match no-credit UL exactly.",
    )
