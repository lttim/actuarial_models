from __future__ import annotations

import pricing_projection as sp
from pricing_ui import _normalize_run_state_for_selected_product
from product_registry import ProductType


def test_spia_switch_forces_default_rp_mode_and_nonblank_paths() -> None:
    state: dict[str, object] = {
        "run_m_mode": "qx_csv",
        "run_rp_xlsx": "",
        "run_rp_out": "   ",
        "run_mp_xlsx": "",
        "run_mp_out": "",
    }
    _normalize_run_state_for_selected_product(
        state,
        selected_product=ProductType.SPIA,
        switched_product=True,
    )
    assert state["run_m_mode"] == "rp2014_mp2016"
    assert state["run_rp_xlsx"] == sp.DEFAULT_RP2014_XLSX
    assert state["run_rp_out"] == sp.DEFAULT_RP2014_MALE_HEALTHY_QX_CSV
    assert state["run_mp_xlsx"] == sp.DEFAULT_MP2016_XLSX
    assert state["run_mp_out"] == sp.DEFAULT_MP2016_MALE_IMPROVEMENT_CSV


def test_invalid_mode_is_normalized_to_selected_product_default() -> None:
    state: dict[str, object] = {
        "run_m_mode": "us_ssa_2015_period",
    }
    _normalize_run_state_for_selected_product(
        state,
        selected_product=ProductType.SPIA,
        switched_product=False,
    )
    assert state["run_m_mode"] == "rp2014_mp2016"


def test_term_product_disables_index_and_mc_toggles() -> None:
    state: dict[str, object] = {
        "run_use_index": True,
        "run_mc_enable": True,
        "run_m_mode": "synthetic",
    }
    _normalize_run_state_for_selected_product(
        state,
        selected_product=ProductType.TERM_LIFE,
        switched_product=False,
    )
    assert state["run_use_index"] is False
    assert state["run_mc_enable"] is False
    assert state["run_m_mode"] in ("us_ssa_2015_period", "qx_csv", "synthetic")


def test_term_product_normalizes_non_positive_monthly_premium_to_default() -> None:
    state: dict[str, object] = {
        "run_term_monthly_premium": 0.0,
    }
    _normalize_run_state_for_selected_product(
        state,
        selected_product=ProductType.TERM_LIFE,
        switched_product=True,
    )
    assert float(state["run_term_monthly_premium"]) == 250.0


def test_term_product_sets_default_monthly_premium_when_missing() -> None:
    state: dict[str, object] = {}
    _normalize_run_state_for_selected_product(
        state,
        selected_product=ProductType.TERM_LIFE,
        switched_product=True,
    )
    assert float(state["run_term_monthly_premium"]) == 250.0


def test_invalid_basic_controls_are_normalized() -> None:
    state: dict[str, object] = {
        "run_sex": "unknown",
        "run_y_mode": "bad",
        "run_expense_mode": "other",
        "run_index_csv": "",
    }
    _normalize_run_state_for_selected_product(
        state,
        selected_product=ProductType.SPIA,
        switched_product=False,
    )
    assert state["run_sex"] == "male"
    assert state["run_y_mode"] == "par_bootstrap"
    assert state["run_expense_mode"] == "csv"
    assert state["run_index_csv"] == sp.DEFAULT_SP500_SCENARIO_CSV
