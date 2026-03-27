from __future__ import annotations

import io

import numpy as np
import pytest
from openpyxl import load_workbook

import pricing_projection as sp
import term_projection as tp
from build_term_excel_workbook import build_term_workbook_from_spec, term_excel_spec_from_launcher


pytestmark = [pytest.mark.parity, pytest.mark.product_term]


def _setup_case() -> tuple[tp.TermLifeContract, sp.YieldCurve, sp.MortalityTableQx]:
    contract = tp.TermLifeContract(issue_age=45, sex="male", death_benefit=250_000.0, monthly_premium=250.0, term_years=20)
    yc = sp.YieldCurve.from_flat_rate(0.04)
    ages = np.arange(0, 121, dtype=int)
    qx = np.full_like(ages, 0.01, dtype=float)
    mort = sp.MortalityTableQx(ages, qx)
    return contract, yc, mort


def test_term_eoy_claims_only_on_policy_year_end_months():
    contract, yc, mort = _setup_case()
    res = tp.price_term_life_level_monthly(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=90,
        spread=0.0,
        valuation_year=None,
    )
    non_zero_months = res.months[np.asarray(res.expected_claim_cashflows) > 0.0]
    assert non_zero_months.size > 0
    assert np.all((non_zero_months % 12) == 0)


def test_term_liability_path_drives_alm_without_shape_drift():
    contract, yc, mort = _setup_case()
    res = tp.price_term_life_level_monthly(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=90,
        spread=0.0,
        valuation_year=None,
    )
    asm = sp.ALMAssumptions(
        allocation=sp.alm_default_allocation_spec(),
        rebalance_band=0.10,
        rebalance_frequency_months=1,
        reinvest_rule="hold_cash",
        disinvest_rule="shortest_first",
        rebalance_policy="liquidity_only",
        liquidity_near_liquid_years=0.25,
    )
    path = tp.liability_path_from_term_projection(res)
    alm = sp.run_alm_projection_from_liability_path(
        liability_path=path,
        yield_curve=yc,
        spread=0.0,
        assumptions=asm,
        initial_asset_market_value=500_000.0,
    )
    assert alm.asset_market_value.shape == res.expected_total_cashflows.shape
    assert alm.liability_pv.shape == res.expected_total_cashflows.shape
    assert np.isfinite(alm.surplus).all()


def test_term_workbook_modelcheck_reconciles_zero_difference():
    contract, yc, mort = _setup_case()
    spec = term_excel_spec_from_launcher(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=90,
        spread=0.0,
        valuation_year=2025,
        expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
        yield_mode_label="flat",
        mortality_mode_label="synthetic",
        expense_mode_label="manual",
    )
    xlsx = build_term_workbook_from_spec(spec)
    wb = load_workbook(io.BytesIO(xlsx), data_only=True)
    ws = wb["ModelCheck"]
    diffs = [float(ws[f"D{r}"].value) for r in (2, 3, 4)]
    np.testing.assert_allclose(np.asarray(diffs, dtype=float), np.zeros(3, dtype=float), atol=0.0, rtol=0.0)
