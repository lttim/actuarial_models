"""Unit tests for :func:`product_registry.validate_run_inputs`.

This pre-flight is the single source of truth for "is this Streamlit
session-state dict launchable into a pricing run". It centralises the
cross-product invariants (issue/horizon age sanity) so the same checks
do not get re-implemented inline in pricing_ui.py.

The test suite covers:

1. The cross-product invariants (issue_age + horizon_age presence and
   ordering).
2. Empty-tuple return for a well-formed minimal input dict.
3. The per-product validator hook is consulted when registered.
"""

from __future__ import annotations

import pytest

from product_registry import (
    ProductType,
    _PRODUCT_VALIDATORS,
    validate_run_inputs,
)

pytestmark = [pytest.mark.invariant]


def test_well_formed_input_yields_no_errors() -> None:
    errors = validate_run_inputs(
        ProductType.SPIA,
        {"issue_age": 65, "horizon_age": 100},
    )
    assert errors == ()


def test_missing_issue_age_is_reported() -> None:
    errors = validate_run_inputs(ProductType.SPIA, {"horizon_age": 100})
    assert any("issue_age is required" in e for e in errors)


def test_missing_horizon_age_is_reported() -> None:
    errors = validate_run_inputs(ProductType.SPIA, {"issue_age": 65})
    assert any("horizon_age is required" in e for e in errors)


def test_horizon_must_exceed_issue() -> None:
    errors = validate_run_inputs(
        ProductType.SPIA,
        {"issue_age": 65, "horizon_age": 65},
    )
    assert any(
        "must be strictly greater than" in e and "issue_age" in e for e in errors
    )


def test_horizon_below_issue_is_reported() -> None:
    errors = validate_run_inputs(
        ProductType.SPIA,
        {"issue_age": 70, "horizon_age": 65},
    )
    assert any("must be strictly greater than" in e for e in errors)


def test_per_product_validator_hook_is_consulted() -> None:
    """Registering a temporary validator under TERM_LIFE must surface its
    error string. Cleanup leaves the registry empty so other tests are
    unaffected."""
    sentinel = "term-only constraint failed: cap < floor"

    def _term_validator(state):
        return [sentinel]

    _PRODUCT_VALIDATORS[ProductType.TERM_LIFE] = _term_validator
    try:
        errors = validate_run_inputs(
            ProductType.TERM_LIFE,
            {"issue_age": 35, "horizon_age": 60},
        )
        assert sentinel in errors
    finally:
        _PRODUCT_VALIDATORS.pop(ProductType.TERM_LIFE, None)


def test_validators_registry_matches_seven_product_rollout() -> None:
    """The seven new products (MYGA / FIA / VA / WL / UL / IUL / VUL)
    register cross-input validators per Section 2 Step G of the
    seven-product rollout. SPIA / Term / RILA do NOT register one; their
    contracts are validated by the engine constructors instead.
    """
    expected_with_validator = {
        ProductType.MYGA,
        ProductType.FIA,
        ProductType.VARIABLE_ANNUITY,
        ProductType.WHOLE_LIFE,
        ProductType.UNIVERSAL_LIFE,
        ProductType.INDEXED_UL,
        ProductType.VARIABLE_UL,
    }
    assert set(_PRODUCT_VALIDATORS) == expected_with_validator, (
        "_PRODUCT_VALIDATORS drifted from the seven-product set. "
        "If you add or remove a per-product validator, update this test "
        "AND add a focused unit test that exercises the new validator."
    )


def test_myga_validator_rejects_zero_premium() -> None:
    errors = validate_run_inputs(
        ProductType.MYGA,
        {
            "issue_age": 60,
            "horizon_age": 70,
            "myga_single_premium": 0.0,
        },
    )
    assert any("myga_single_premium" in e for e in errors)


def test_fia_validator_rejects_cap_below_floor() -> None:
    errors = validate_run_inputs(
        ProductType.FIA,
        {
            "issue_age": 60,
            "horizon_age": 70,
            "fia_cap": 0.0,
            "fia_floor": 0.05,
        },
    )
    assert any("must be >= floor" in e for e in errors)


def test_va_validator_rejects_excessive_me() -> None:
    errors = validate_run_inputs(
        ProductType.VARIABLE_ANNUITY,
        {
            "issue_age": 55,
            "horizon_age": 75,
            "va_me_charge_annual": 0.20,
        },
    )
    assert any("va_me_charge_annual" in e for e in errors)


def test_wl_validator_rejects_negative_face() -> None:
    errors = validate_run_inputs(
        ProductType.WHOLE_LIFE,
        {
            "issue_age": 45,
            "horizon_age": 120,
            "wl_face_amount": -10_000.0,
        },
    )
    assert any("wl_face_amount" in e for e in errors)


def test_ul_validator_rejects_invalid_load() -> None:
    errors = validate_run_inputs(
        ProductType.UNIVERSAL_LIFE,
        {
            "issue_age": 45,
            "horizon_age": 120,
            "ul_face_amount": 250_000.0,
            "ul_single_premium": 25_000.0,
            "ul_premium_load_pct": 1.5,
        },
    )
    assert any("ul_premium_load_pct" in e for e in errors)


def test_iul_validator_rejects_cap_below_floor() -> None:
    errors = validate_run_inputs(
        ProductType.INDEXED_UL,
        {
            "issue_age": 45,
            "horizon_age": 120,
            "iul_cap": 0.0,
            "iul_floor": 0.05,
        },
    )
    assert any("must be >= floor" in e for e in errors)


def test_vul_validator_rejects_zero_face() -> None:
    errors = validate_run_inputs(
        ProductType.VARIABLE_UL,
        {
            "issue_age": 45,
            "horizon_age": 120,
            "vul_face_amount": 0.0,
        },
    )
    assert any("vul_face_amount" in e for e in errors)
