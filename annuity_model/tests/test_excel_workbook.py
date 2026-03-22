"""Smoke tests for Excel workbook builder."""

from __future__ import annotations

import io

import numpy as np
import pytest
from openpyxl import load_workbook

import spia_projection as sp
from build_spia_excel_workbook import (
    ExcelPythonSnapshot,
    build_workbook_from_spec,
    excel_spec_from_launcher,
)


def test_model_check_sheet_embeds_python_snapshot():
    contract = sp.SPIAContract(issue_age=65, sex="male", benefit_annual=100_000.0)
    yc = sp.YieldCurve.from_flat_rate(0.04)
    ages = np.arange(0, 121, dtype=int)
    qx = np.full_like(ages, 0.02, dtype=float)
    mort = sp.MortalityTableQx(ages, qx)
    ex = sp.ExpenseAssumptions(0.0, 0.0, 0.0)
    res = sp.price_spia_single_premium(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=80,
        spread=0.0,
        valuation_year=None,
        expenses=ex,
        index_scenario_csv_path=None,
        expense_annual_inflation=0.0,
    )
    spec = excel_spec_from_launcher(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=80,
        spread=0.0,
        valuation_year=2025,
        expenses=ex,
        yield_mode_label="flat",
        mortality_mode_label="synthetic",
        expense_mode_label="manual",
        index_s0=float(res.index_s0),
        index_levels_at_payment=res.index_level_at_payment,
        expense_annual_inflation=0.0,
    )
    snap = ExcelPythonSnapshot(
        pv_benefit=float(res.pv_benefit),
        pv_monthly_expenses=float(res.pv_monthly_expenses),
        pv_monthly_total=float(res.pv_benefit + res.pv_monthly_expenses),
        single_premium=float(res.single_premium),
        annuity_factor=float(res.annuity_factor),
    )
    raw = build_workbook_from_spec(spec, out_path=None, python_snapshot=snap)
    wb = load_workbook(io.BytesIO(raw))
    assert "ModelCheck" in wb.sheetnames
    mc = wb["ModelCheck"]
    assert mc["B5"].value == pytest.approx(res.pv_benefit, rel=1e-9)
    assert mc["B9"].value == pytest.approx(res.annuity_factor, rel=1e-9)
    assert mc["C5"].value == "=Projection!X4"
