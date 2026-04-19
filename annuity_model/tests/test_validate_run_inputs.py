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


def test_validators_registry_empty_by_default() -> None:
    """No product registers an extra validator yet (P1 ships the hook
    without enabling it). When a product opts in, this test is the place
    to assert the new entry is intentional."""
    assert _PRODUCT_VALIDATORS == {}, (
        "_PRODUCT_VALIDATORS gained an entry. Update this test to assert "
        "the new entry is intentional, and add a per-product test that "
        "exercises the registered validator."
    )
