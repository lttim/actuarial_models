from __future__ import annotations

import numpy as np

import pricing_projection as sp
from product_registry import (
    ProductType,
    get_product_capabilities,
    get_product_default_mortality_mode,
    get_product_mortality_mode_options,
)
from pricing_ui import _build_mortality


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

