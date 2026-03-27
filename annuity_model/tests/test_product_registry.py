from __future__ import annotations

import numpy as np
import pytest

import pricing_projection as sp
import term_projection as tp
from build_pricing_excel_workbook import excel_spec_from_launcher
from product_registry import (
    ProductType,
    get_mortality_mode_label,
    get_pricing_metrics,
    get_product_adapter,
    get_product_ui_config,
    get_term_contract_ui_config,
)


pytestmark = [pytest.mark.product_spia, pytest.mark.product_term]


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


def test_term_adapter_price_and_dispatch():
    contract, yc, mort, ex = _setup_case()
    term_contract = tp.TermLifeContract(
        issue_age=contract.issue_age,
        sex=contract.sex,
        death_benefit=250_000.0,
        monthly_premium=250.0,
        term_years=20,
    )
    adapter = get_product_adapter(ProductType.TERM_LIFE)
    res = adapter.price(
        contract=term_contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=95,
        spread=0.0,
        valuation_year=None,
        expenses=ex,
        expenses_csv_path=sp.DEFAULT_EXPENSES_CSV,
        index_scenario_csv_path=None,
        expense_annual_inflation=0.0,
    )
    direct = tp.price_term_life_level_monthly(
        contract=term_contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=95,
        spread=0.0,
        valuation_year=None,
    )
    np.testing.assert_allclose(res.expected_total_cashflows, direct.expected_total_cashflows, rtol=0, atol=0)
    assert float(res.pv_benefit) == pytest.approx(float(direct.pv_benefit), rel=0, abs=0)


def test_term_adapter_monte_carlo_not_available():
    contract, yc, mort, ex = _setup_case()
    term_contract = tp.TermLifeContract(issue_age=contract.issue_age, sex=contract.sex, death_benefit=250_000.0)
    adapter = get_product_adapter(ProductType.TERM_LIFE)
    with pytest.raises(NotImplementedError):
        adapter.price_monte_carlo(
            contract=term_contract,
            yield_curve=yc,
            mortality=mort,
            horizon_age=95,
            spread=0.0,
            valuation_year=None,
            expenses=ex,
            expenses_csv_path=sp.DEFAULT_EXPENSES_CSV,
            expense_annual_inflation=0.0,
            n_sims=100,
            annual_drift=0.06,
            annual_vol=0.15,
            seed=42,
            s0=100.0,
        )


def test_term_contract_ui_config_defaults():
    cfg = get_term_contract_ui_config()
    assert cfg.death_benefit_label == "Death benefit ($)"
    assert cfg.default_death_benefit == pytest.approx(250_000.0, rel=0, abs=0)
    assert cfg.term_length_options == ("20 years",)
    assert cfg.premium_mode_options == ("Level monthly",)
    assert cfg.benefit_timing_options == ("EOY death benefit",)
    assert cfg.default_monthly_premium == pytest.approx(250.0, rel=0, abs=0)


def test_mortality_mode_label_helper_has_expected_values():
    assert get_mortality_mode_label("synthetic") == "Synthetic (demo, wide age range)"
    assert get_mortality_mode_label("us_ssa_2015_period").startswith("US SSA 2015 period life table")
    assert get_mortality_mode_label("unknown_mode_key") == "unknown_mode_key"


def test_pricing_metrics_for_spia_product():
    contract, yc, mort, ex = _setup_case()
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
    metrics = get_pricing_metrics(ProductType.SPIA, res)
    assert tuple(m.label for m in metrics) == ("Single premium", "PV benefit", "PV monthly expenses", "Annuity factor")
    assert tuple(m.is_money for m in metrics) == (True, True, True, False)
    assert metrics[0].value == pytest.approx(float(res.single_premium), rel=0, abs=0)


def test_pricing_metrics_for_term_product():
    contract, yc, mort, _ = _setup_case()
    term_contract = tp.TermLifeContract(
        issue_age=contract.issue_age,
        sex=contract.sex,
        death_benefit=250_000.0,
        monthly_premium=250.0,
        term_years=20,
    )
    res = tp.price_term_life_level_monthly(
        contract=term_contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=95,
        spread=0.0,
        valuation_year=None,
    )
    metrics = get_pricing_metrics(ProductType.TERM_LIFE, res)
    assert tuple(m.label for m in metrics) == (
        "PV claims",
        "PV premiums",
        "Net PV (claims - premiums)",
        "Issue reserve",
    )
    assert all(m.is_money for m in metrics)
    assert metrics[0].value == pytest.approx(float(res.pv_benefit), rel=0, abs=0)
    assert metrics[1].value == pytest.approx(float(-res.pv_monthly_expenses), rel=0, abs=0)


def test_product_ui_config_export_filenames():
    spia_ui_cfg = get_product_ui_config(ProductType.SPIA)
    term_ui_cfg = get_product_ui_config(ProductType.TERM_LIFE)
    assert spia_ui_cfg.projection_csv_filename == "pricing_projection_spia.csv"
    assert term_ui_cfg.projection_csv_filename == "pricing_projection_term_life.csv"
    assert spia_ui_cfg.recalc_workbook_filename == "spia_recalc_model.xlsx"
    assert term_ui_cfg.recalc_workbook_filename == "term_life_recalc_model.xlsx"
