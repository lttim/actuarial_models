from __future__ import annotations

import numpy as np
import pytest

import spia_projection as sp
from build_spia_excel_workbook import excel_spec_from_launcher
from product_registry import ProductType, get_product_adapter


pytestmark = [pytest.mark.product_spia]


def _setup_case() -> tuple[sp.SPIAContract, sp.YieldCurve, sp.MortalityTableQx, sp.ExpenseAssumptions]:
    contract = sp.SPIAContract(issue_age=65, sex="male", benefit_annual=100_000.0)
    yc = sp.YieldCurve.from_flat_rate(0.04)
    ages = np.arange(0, 121, dtype=int)
    qx = np.full_like(ages, 0.02, dtype=float)
    mort = sp.MortalityTableQx(ages, qx)
    ex = sp.ExpenseAssumptions(0.0, 0.0, 0.0)
    return contract, yc, mort, ex


def test_spia_adapter_price_matches_legacy():
    contract, yc, mort, ex = _setup_case()
    adapter = get_product_adapter(ProductType.SPIA)
    res_adapter = adapter.price(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=80,
        spread=0.0,
        valuation_year=None,
        expenses=ex,
        expenses_csv_path=sp.DEFAULT_EXPENSES_CSV,
        index_scenario_csv_path=None,
        expense_annual_inflation=0.0,
    )
    res_legacy = sp.price_spia_single_premium(
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
    assert isinstance(res_adapter, sp.SPIAProjectionResult)
    np.testing.assert_allclose(res_adapter.expected_total_cashflows, res_legacy.expected_total_cashflows, rtol=0, atol=0)
    assert float(res_adapter.single_premium) == pytest.approx(float(res_legacy.single_premium), rel=0, abs=0)


def test_spia_alm_generic_entrypoint_matches_legacy_wrapper():
    contract, yc, mort, ex = _setup_case()
    pricing = sp.price_spia_single_premium(
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
    asm = sp.ALMAssumptions(
        allocation=sp.alm_default_allocation_spec(),
        rebalance_band=0.10,
        rebalance_frequency_months=1,
        reinvest_rule="hold_cash",
        disinvest_rule="shortest_first",
        rebalance_policy="liquidity_only",
        liquidity_near_liquid_years=0.25,
    )
    legacy = sp.run_alm_projection(
        pricing=pricing,
        yield_curve=yc,
        spread=0.0,
        assumptions=asm,
        initial_asset_market_value=float(pricing.single_premium),
    )
    generic = sp.run_alm_projection_from_liability_path(
        liability_path=sp.liability_path_from_spia_projection(pricing),
        yield_curve=yc,
        spread=0.0,
        assumptions=asm,
        initial_asset_market_value=float(pricing.single_premium),
    )
    np.testing.assert_allclose(generic.asset_market_value, legacy.asset_market_value, rtol=0, atol=0)
    np.testing.assert_allclose(generic.liability_pv, legacy.liability_pv, rtol=0, atol=0)
    np.testing.assert_allclose(generic.surplus, legacy.surplus, rtol=0, atol=0)


def test_spia_adapter_excel_spec_matches_legacy_builder():
    contract, yc, mort, ex = _setup_case()
    pricing = sp.price_spia_single_premium(
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
    adapter = get_product_adapter(ProductType.SPIA)
    spec_adapter = adapter.excel_spec_from_run(
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
        index_s0=float(pricing.index_s0),
        index_levels_at_payment=pricing.index_level_at_payment,
        expense_annual_inflation=0.0,
    )
    spec_legacy = excel_spec_from_launcher(
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
        index_s0=float(pricing.index_s0),
        index_levels_at_payment=pricing.index_level_at_payment,
        expense_annual_inflation=0.0,
    )
    assert spec_adapter.n_months == spec_legacy.n_months
    assert spec_adapter.issue_age == spec_legacy.issue_age
    assert spec_adapter.benefit_annual == pytest.approx(spec_legacy.benefit_annual, rel=0, abs=0)


def test_unimplemented_product_types_raise():
    with pytest.raises(NotImplementedError):
        get_product_adapter(ProductType.WHOLE_LIFE)
    with pytest.raises(NotImplementedError):
        get_product_adapter(ProductType.VARIABLE_ANNUITY)
