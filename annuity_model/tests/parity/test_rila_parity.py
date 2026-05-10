from __future__ import annotations

import io

import numpy as np
import pytest
from openpyxl import load_workbook

from annuity_model import pricing_projection as sp
from annuity_model import rila_projection as rp
from annuity_model.build_pricing_excel_workbook import LIABILITY_SHEET_NAME
from annuity_model.build_rila_excel_workbook import (
    build_rila_workbook_from_spec,
    rila_excel_spec_from_launcher,
)
from annuity_model.parity_constants import RILA_AV_TOL, RILA_PV_TOL

pytestmark = [pytest.mark.parity, pytest.mark.product_rila]


def _setup_case() -> tuple[
    rp.RILAContract, sp.YieldCurve, sp.MortalityTableQx, sp.ExpenseAssumptions
]:
    contract = rp.RILAContract(
        issue_age=55,
        sex="male",
        participation=0.85,
        cap=0.09,
        floor=-0.02,
        rider_fee_annual=0.008,
    )
    yc = sp.YieldCurve.from_flat_rate(0.035)
    ages = np.arange(0, 121, dtype=int)
    qx = np.clip(0.008 + ages * 2e-5, 1e-6, 0.4)
    mort = sp.MortalityTableQx(ages, qx)
    ex = sp.ExpenseAssumptions(0.0, 0.0, 25.0)
    return contract, yc, mort, ex


def test_rila_workbook_modelcheck_reconciles_zero_difference():
    contract, yc, mort, ex = _setup_case()
    n_months = int(round((90 - contract.issue_age) * 12))
    rng = np.random.default_rng(42)
    levels = 100.0 * np.cumprod(1.0 + rng.normal(0.004, 0.02, size=n_months))
    res = rp.price_rila_single_premium(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=90,
        spread=0.0,
        valuation_year=None,
        expenses=ex,
        index_s0=100.0,
        index_levels_payment=levels,
        expense_annual_inflation=0.01,
    )
    spec = rila_excel_spec_from_launcher(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=90,
        spread=0.0,
        valuation_year=2025,
        expenses=ex,
        yield_mode_label="flat",
        mortality_mode_label="synthetic",
        expense_mode_label="manual",
        index_s0=float(res.index_s0),
        index_levels_at_payment=res.index_level_at_payment,
        expense_annual_inflation=0.01,
    )
    xlsx = build_rila_workbook_from_spec(spec)
    wb = load_workbook(io.BytesIO(xlsx), data_only=False)
    ws_mc = wb["ModelCheck"]
    ws_liab = wb[LIABILITY_SHEET_NAME]
    assert ws_mc["C5"].value == f"={LIABILITY_SHEET_NAME}!X4"
    assert ws_mc["C6"].value == f"={LIABILITY_SHEET_NAME}!X5"
    assert ws_mc["C7"].value == f"={LIABILITY_SHEET_NAME}!X7"
    assert ws_mc["C8"].value == f"={LIABILITY_SHEET_NAME}!X8"
    for coord, needle in (
        ("A4", "=IF(ROW()-3>"),
        ("J4", "=IF(A4="),
        ("H4", "=IF(A4="),
        ("O4", "=IF(A4="),
        ("P4", "=IF(A4="),
    ):
        v = ws_liab[coord].value
        assert isinstance(v, str) and v.startswith("="), coord
        assert needle in v, (coord, v)

    pv_b = float(np.sum(res.expected_benefit_cashflows * res.discount_factors))
    pv_e = float(np.sum(res.expected_expense_cashflows * res.discount_factors))
    pv_t = float(np.sum(res.expected_total_cashflows * res.discount_factors))
    # PV / ModelCheck tolerance comes from parity_constants.RILA_PV_TOL.
    np.testing.assert_allclose(float(ws_mc["B5"].value), pv_b, rtol=0.0, atol=RILA_PV_TOL)
    np.testing.assert_allclose(float(ws_mc["B6"].value), pv_e, rtol=0.0, atol=RILA_PV_TOL)
    np.testing.assert_allclose(float(ws_mc["B7"].value), pv_t, rtol=0.0, atol=RILA_PV_TOL)
    np.testing.assert_allclose(
        float(ws_mc["B8"].value), float(res.single_premium), rtol=0.0, atol=RILA_PV_TOL
    )


def test_rila_account_value_matches_formula_track_month_12():
    contract, yc, mort, ex = _setup_case()
    levels = np.array([100.0 * (1.01**i) for i in range(120)], dtype=float)
    res = rp.price_rila_single_premium(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=65,
        spread=0.0,
        valuation_year=None,
        expenses=ex,
        index_s0=100.0,
        index_levels_payment=levels,
        expense_annual_inflation=0.0,
    )
    av12 = float(res.account_value_end_month[11])
    L = rp.levels_end_by_policy_month(s0=100.0, levels_payment=levels)
    raw = L[12] / L[0] - 1.0
    cr = rp.segment_credited_return(
        raw=raw,
        participation=contract.participation,
        cap=contract.cap,
        floor=contract.floor,
    )
    prem = float(res.single_premium)
    expect = prem * (1.0 + cr) * (1.0 - contract.rider_fee_annual / 12.0) ** 12
    np.testing.assert_allclose(av12, expect, rtol=0.0, atol=RILA_AV_TOL)
