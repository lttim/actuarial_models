"""Unit tests for pricing run-form defaults and coercion (no Streamlit runtime)."""

from __future__ import annotations

from pricing_run_form_state import (
    build_run_form_seed_defaults,
    coerce_numeric_widget_value,
    ensure_session_choice,
)
from product_registry import ProductType


def test_coerce_numeric_clamps_and_replaces_non_positive() -> None:
    assert coerce_numeric_widget_value(None, 2025, min_value=1950, max_value=2100) == 2025
    assert coerce_numeric_widget_value(2014, 2025, min_value=1950, max_value=2100) == 2014
    assert coerce_numeric_widget_value("bad", 2025, min_value=1950, max_value=2100) == 2025
    assert coerce_numeric_widget_value(1800, 2025, min_value=1950, max_value=2100) == 1950
    assert coerce_numeric_widget_value(3000, 2025, min_value=1950, max_value=2100) == 2100
    assert coerce_numeric_widget_value(0.0, 250.0, min_value=0.0, replace_non_positive=True) == 250.0
    assert coerce_numeric_widget_value(100.0, 250.0, min_value=0.0, replace_non_positive=True) == 100.0


def test_build_run_form_seed_defaults_matches_expected_keys() -> None:
    d = build_run_form_seed_defaults(
        product_default=ProductType.SPIA.value,
        saved_inputs={},
        meta={},
        default_product_type=ProductType.SPIA,
    )
    assert d["run_valuation_year"] == 2025
    assert d["run_horizon_age"] == 110
    assert d["run_m_mode"] == "rp2014_mp2016"
    assert d["run_spia_benefit_annual"] == 100_000.0
    assert "run_flat_rate" in d
    assert "run_term_monthly_premium" in d


def test_ensure_session_choice_fixes_invalid_option() -> None:
    state: dict[str, object] = {"run_m_mode": "us_ssa_2015_period"}
    allowed = ("synthetic", "qx_csv", "rp2014_mp2016")
    ensure_session_choice(state, "run_m_mode", allowed, "rp2014_mp2016")
    assert state["run_m_mode"] == "rp2014_mp2016"
