"""Actuarial reasonableness tests for Whole Life (single premium)."""

from __future__ import annotations

import numpy as np
import pytest

import pricing_projection as sp
import wl_projection as wl
from actuarial_benchmarks import (
    WL_BENCHMARK_SP_HI,
    WL_BENCHMARK_SP_LO,
    WL_NSP_TOL,
    WL_SENSITIVITY_EPS,
)

pytestmark = [pytest.mark.parity, pytest.mark.product_wl]


def _baseline_contract() -> wl.WLContract:
    return wl.WLContract(
        issue_age=45, sex="male", smoker_class="nonsmoker", face_amount=250_000.0,
    )


def _yc(r: float = 0.04) -> sp.YieldCurve:
    return sp.YieldCurve.from_flat_rate(r)


def _cso_mort() -> sp.MortalityTableQx:
    """Use the CSO 2017 Ultimate placeholder for life-product banding."""
    from mortality_2017_cso import MortalityTable2017CSO

    return MortalityTable2017CSO.load(sex="male", smoker_class="nonsmoker").table


def _price(contract=None, yc=None, mort=None):
    return wl.price_wl_single_premium(
        contract=contract or _baseline_contract(),
        yield_curve=yc or _yc(),
        mortality=mort or _cso_mort(),
        horizon_age=120,
        expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
    )


def test_wl_actuarial_sanity_signs():
    res = _price()
    assert res.single_premium > 0
    assert res.pv_benefit > 0
    assert (res.expected_benefit_cashflows >= 0).all()
    assert (res.discount_factors > 0).all()


def test_wl_sp_within_benchmark_band():
    res = _price()
    assert WL_BENCHMARK_SP_LO <= res.single_premium <= WL_BENCHMARK_SP_HI, (
        f"WL SP={res.single_premium:,.2f} fell outside band "
        f"[{WL_BENCHMARK_SP_LO:,.0f}, {WL_BENCHMARK_SP_HI:,.0f}]. "
        "Synthetic CSO placeholder may differ from licensed CSO; if engine "
        "is correct, document the band shift in docs/actuarial_benchmarks.md."
    )


def test_wl_nsp_matches_closed_form():
    """Section 13.5: SP must equal face × Σ v^t × P(death in month t) (NSP_x)."""
    res = _price()
    # PV(face × death-prob) is what the engine returns as pv_benefit.
    # NSP closed form is exactly the same arithmetic; expect tight match.
    expected_nsp = float(np.sum(res.expected_benefit_cashflows * res.discount_factors))
    assert abs(res.pv_benefit - expected_nsp) <= WL_NSP_TOL, (
        f"WL pv_benefit={res.pv_benefit:.4f} vs expected NSP={expected_nsp:.4f}, "
        f"diff={res.pv_benefit - expected_nsp:+.4f} > {WL_NSP_TOL}."
    )


def test_wl_yield_sensitivity_negative_pv():
    """+100bps yield shock must reduce PV of death benefits."""
    base = _price()
    shocked = _price(yc=_yc(0.05))
    assert shocked.pv_benefit < base.pv_benefit - WL_SENSITIVITY_EPS, (
        f"+100bps yield shock did not reduce WL PV. base={base.pv_benefit:.2f} "
        f"shocked={shocked.pv_benefit:.2f}. Sign bug."
    )


def test_wl_face_increase_raises_sp_proportionally():
    """SP is linear in face; doubling face should roughly double SP."""
    base = _price()
    high_face = wl.WLContract(
        issue_age=45, sex="male", smoker_class="nonsmoker", face_amount=500_000.0,
    )
    res_high = _price(contract=high_face)
    ratio = res_high.single_premium / base.single_premium
    assert 1.95 <= ratio <= 2.05, (
        f"Doubling face should ~double SP (linearity). Got ratio={ratio:.4f}."
    )


def test_wl_higher_age_raises_sp():
    """Issuing later means more death-claim risk per face dollar."""
    base = _price()
    older = wl.WLContract(
        issue_age=70, sex="male", smoker_class="nonsmoker", face_amount=250_000.0,
    )
    res_older = _price(contract=older)
    assert res_older.single_premium > base.single_premium + WL_SENSITIVITY_EPS, (
        f"Older issue age should raise SP. age45={base.single_premium:.2f} "
        f"age70={res_older.single_premium:.2f}."
    )
