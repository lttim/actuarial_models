from __future__ import annotations

import ast
from pathlib import Path

import numpy as np
import pytest

from annuity_model import pricing_projection as sp
from annuity_model import term_projection as tp
from annuity_model.pricing_ui import _build_mortality
from annuity_model.product_registry import (
    ProductType,
    get_product_capabilities,
    get_product_default_mortality_mode,
    get_product_mortality_mode_options,
    get_term_contract_ui_config,
    parse_term_benefit_timing_label,
    parse_term_length_label_to_years,
    parse_term_premium_mode_label,
    term_benefit_timing_label_options,
    term_length_label_options,
    term_premium_mode_label_options,
)


def test_term_capabilities_disable_scenario_and_mc() -> None:
    caps = get_product_capabilities(ProductType.TERM_LIFE)
    assert caps.supports_economic_scenario is False
    assert caps.supports_monte_carlo is False


def test_term_mortality_mode_helper_defaults_to_ssa() -> None:
    options = get_product_mortality_mode_options(ProductType.TERM_LIFE)
    default_mode = get_product_default_mortality_mode(ProductType.TERM_LIFE)
    assert default_mode == "us_ssa_2015_period"
    assert default_mode in options


def test_term_default_mortality_uses_ssa_sex_specific_qx() -> None:
    male, needs_vy_m = _build_mortality(
        "us_ssa_2015_period",
        product_type=ProductType.TERM_LIFE,
        sex="male",
        qx_csv="unused.csv",
        rp_xlsx="unused.xlsx",
        rp_out_csv="unused.csv",
        mp_xlsx="unused.xlsx",
        mp_out_csv="unused.csv",
    )
    female, needs_vy_f = _build_mortality(
        "us_ssa_2015_period",
        product_type=ProductType.TERM_LIFE,
        sex="female",
        qx_csv="unused.csv",
        rp_xlsx="unused.xlsx",
        rp_out_csv="unused.csv",
        mp_xlsx="unused.xlsx",
        mp_out_csv="unused.csv",
    )
    assert isinstance(male, sp.MortalityTableQx)
    assert isinstance(female, sp.MortalityTableQx)
    assert needs_vy_m is False
    assert needs_vy_f is False
    assert male.qx_at_int_age(65) > female.qx_at_int_age(65)
    assert np.isclose(float(male.qx_at_int_age(65)), 0.015967, atol=1e-12)
    assert np.isclose(float(female.qx_at_int_age(65)), 0.009794, atol=1e-12)


# ---------------------------------------------------------------------------
# Term widget → engine-contract wiring
#
# Regression: 2026-04-18. pricing_ui.py used to hard-code term_years=20,
# premium_mode="level_monthly", benefit_timing="eoy_death" when constructing
# TermLifeContract, silently dropping the widget choices. The fix wires every
# selectbox value through product_registry.parse_term_*_label_to_value and
# the tests below ensure it never regresses.
# ---------------------------------------------------------------------------


def test_term_ui_config_options_match_parser_maps() -> None:
    """The Streamlit selectbox option tuples MUST equal the parser maps.

    If a developer adds an option label to TermContractUIConfig.*_options
    without adding the matching engine-value mapping, the UI will raise at
    runtime when the user selects it. This test makes that drift a CI failure
    instead of a runtime crash.
    """
    cfg = get_term_contract_ui_config()
    assert tuple(cfg.term_length_options) == term_length_label_options()
    assert tuple(cfg.premium_mode_options) == term_premium_mode_label_options()
    assert tuple(cfg.benefit_timing_options) == term_benefit_timing_label_options()


def test_every_term_widget_label_parses_to_engine_value() -> None:
    cfg = get_term_contract_ui_config()
    for label in cfg.term_length_options:
        years = parse_term_length_label_to_years(label)
        assert isinstance(years, int) and years > 0
    for label in cfg.premium_mode_options:
        mode = parse_term_premium_mode_label(label)
        assert mode in {"level_monthly"}
    for label in cfg.benefit_timing_options:
        timing = parse_term_benefit_timing_label(label)
        assert timing in {"eoy_death"}


def test_unknown_term_labels_raise_with_helpful_message() -> None:
    with pytest.raises(ValueError, match="Unknown Term length label"):
        parse_term_length_label_to_years("99 years")
    with pytest.raises(ValueError, match="Unknown Term premium mode label"):
        parse_term_premium_mode_label("Annual lump sum")
    with pytest.raises(ValueError, match="Unknown Term benefit timing label"):
        parse_term_benefit_timing_label("Immediate")


def test_parsed_labels_round_trip_into_term_life_contract() -> None:
    cfg = get_term_contract_ui_config()
    contract = tp.TermLifeContract(
        issue_age=65,
        sex="male",
        death_benefit=250_000.0,
        monthly_premium=200.0,
        term_years=parse_term_length_label_to_years(cfg.term_length_options[0]),
        premium_mode=parse_term_premium_mode_label(cfg.premium_mode_options[0]),  # type: ignore[arg-type]
        benefit_timing=parse_term_benefit_timing_label(  # type: ignore[arg-type]
            cfg.benefit_timing_options[0]
        ),
    )
    assert contract.term_years == 20
    assert contract.premium_mode == "level_monthly"
    assert contract.benefit_timing == "eoy_death"


# ---------------------------------------------------------------------------
# AST guard: pricing_ui must not hard-code Term contract literals.
# ---------------------------------------------------------------------------


def _find_term_life_contract_calls(tree: ast.AST) -> list[ast.Call]:
    calls: list[ast.Call] = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Call):
            func = node.func
            attr = getattr(func, "attr", None)
            name = getattr(func, "id", None)
            if attr == "TermLifeContract" or name == "TermLifeContract":
                calls.append(node)
    return calls


def test_pricing_ui_does_not_hardcode_term_contract_fields() -> None:
    """The Term contract built inside pricing_ui MUST source term_years,
    premium_mode, and benefit_timing from parsed widget values, not literal
    constants. We assert via AST that no kwarg in the TermLifeContract call
    inside pricing_ui.py is a Constant for any of those fields. This catches
    the original 2026-04 regression and any future repeats."""
    pricing_ui_path = (
        Path(__file__).resolve().parent.parent / "src" / "annuity_model" / "pricing_ui.py"
    )
    tree = ast.parse(pricing_ui_path.read_text(encoding="utf-8"))
    calls = _find_term_life_contract_calls(tree)
    assert calls, "Expected at least one TermLifeContract call in pricing_ui.py"
    forbidden_constant_kwargs = {"term_years", "premium_mode", "benefit_timing"}
    for call in calls:
        for kw in call.keywords:
            if kw.arg in forbidden_constant_kwargs and isinstance(kw.value, ast.Constant):
                raise AssertionError(
                    f"pricing_ui.py hard-codes TermLifeContract.{kw.arg}="
                    f"{kw.value.value!r}; route the widget choice through "
                    f"product_registry.parse_term_*_label_to_value instead."
                )
