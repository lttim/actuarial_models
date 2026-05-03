"""Actuarial reasonableness tests for Variable UL."""

from __future__ import annotations

import numpy as np
import pytest

import pricing_projection as sp
import ul_projection as ul
import vul_projection as vul
from actuarial_benchmarks import (
    VUL_BENCHMARK_AV_20Y_MC_HI,
    VUL_BENCHMARK_AV_20Y_MC_LO,
    VUL_SENSITIVITY_EPS,
)
from parity_constants import AV_TOL

pytestmark = [pytest.mark.parity, pytest.mark.product_vul]


def _baseline_contract() -> vul.VULContract:
    return vul.VULContract(
        issue_age=45,
        sex="male",
        face_amount=250_000.0,
        single_premium=25_000.0,
        premium_load_pct=0.06,
        monthly_expense_charge=7.50,
    )


def _yc(r: float = 0.04) -> sp.YieldCurve:
    return sp.YieldCurve.from_flat_rate(r)


def _mort() -> sp.MortalityTableQx:
    ages = np.arange(0, 121, dtype=int)
    qx = np.clip(0.005 + ages * 1e-5, 1e-6, 0.4)
    return sp.MortalityTableQx(ages, qx)


def test_vul_actuarial_sanity_signs():
    contract = _baseline_contract()
    n = (120 - 45) * 12
    rng = np.random.default_rng(42)
    levels = 100.0 * np.cumprod(1.0 + rng.normal(0.005, 0.03, size=n))
    res = vul.price_vul_single_premium(
        contract=contract,
        yield_curve=_yc(),
        mortality=_mort(),
        horizon_age=120,
        expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
        index_s0=100.0,
        index_levels_payment=levels,
    )
    assert res.single_premium > 0
    assert res.pv_benefit >= 0
    assert (res.account_value_end_month >= 0).all()


def test_vul_collapses_to_ul_with_zero_subaccount_return():
    """Section 13.5: VUL with constant index = UL with declared_rate=0."""
    contract_vul = _baseline_contract()
    contract_ul = ul.ULContract(
        issue_age=45,
        sex="male",
        face_amount=250_000.0,
        single_premium=25_000.0,
        premium_load_pct=0.06,
        monthly_expense_charge=7.50,
        declared_rate_annual=0.0,
    )
    n = (80 - 45) * 12
    flat_levels = np.full(n, 100.0)
    res_vul = vul.price_vul_single_premium(
        contract=contract_vul,
        yield_curve=_yc(),
        mortality=_mort(),
        horizon_age=80,
        expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
        index_s0=100.0,
        index_levels_payment=flat_levels,
    )
    res_ul = ul.price_ul_single_premium(
        contract=contract_ul,
        yield_curve=_yc(),
        mortality=_mort(),
        horizon_age=80,
        expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
    )
    np.testing.assert_allclose(
        res_vul.account_value_end_month,
        res_ul.account_value_end_month,
        rtol=0.0,
        atol=AV_TOL,
        err_msg="VUL with flat index should match no-credit UL exactly.",
    )


def test_vul_mc_mean_within_band():
    """Section 13.3 MC band $5k-$250k for 20y, 6% drift, 15% vol."""
    contract = _baseline_contract()
    res = vul.price_vul_single_premium_monte_carlo(
        contract=contract,
        yield_curve=_yc(),
        mortality=_mort(),
        horizon_age=80,
        expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
        n_sims=40,
        annual_drift=0.06,
        annual_vol=0.15,
        seed=42,
        s0=100.0,
    )
    # We use 20-y horizon. Use res.av_end_mean as the 20y mean.
    # Note: smoke check; actual MC bands accept wide range.
    assert VUL_BENCHMARK_AV_20Y_MC_LO <= res.av_end_mean <= VUL_BENCHMARK_AV_20Y_MC_HI, (
        f"VUL E[AV(20y)] MC mean={res.av_end_mean:,.2f} fell outside band "
        f"[{VUL_BENCHMARK_AV_20Y_MC_LO:,.0f}, {VUL_BENCHMARK_AV_20Y_MC_HI:,.0f}]."
    )


def test_vul_drift_increase_raises_av():
    """Section 13.4: higher sub-account drift -> higher expected AV."""
    contract = _baseline_contract()
    base = vul.price_vul_single_premium_monte_carlo(
        contract=contract,
        yield_curve=_yc(),
        mortality=_mort(),
        horizon_age=80,
        expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
        n_sims=20,
        annual_drift=0.04,
        annual_vol=0.15,
        seed=42,
        s0=100.0,
    )
    higher = vul.price_vul_single_premium_monte_carlo(
        contract=contract,
        yield_curve=_yc(),
        mortality=_mort(),
        horizon_age=80,
        expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
        n_sims=20,
        annual_drift=0.10,
        annual_vol=0.15,
        seed=42,
        s0=100.0,
    )
    assert higher.av_end_mean > base.av_end_mean + VUL_SENSITIVITY_EPS, (
        f"Higher drift should raise E[AV]. base={base.av_end_mean:.2f} "
        f"higher={higher.av_end_mean:.2f}."
    )
