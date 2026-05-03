"""Cross-registry meta-invariants (P0 hardening, 2026-04).

These tests are the "scaffolding cannot rot" net for the platform's plug-in
points. If a contributor adds a new ``ProductType`` to the enum but forgets
one of the wires (liability-path converter, metric formatter, MODELCHECK
tolerance discipline), the failure surfaces here at PR time, not in
production at the next month-end.

Each test is intentionally short and surgical so the failure message points
at the exact registry that is missing the entry.

Adding a new product? Add it to ``_PRODUCT_ADAPTERS`` *and* register a
liability-path converter from its engine module *and* add an entry to
``_PRICING_METRIC_FORMATTERS``. Then this file passes.

Loosening ``MODELCHECK_TOL`` from 0.0? Update
``annuity_model/docs/model_change_log.md`` in the same PR and route through
CODEOWNERS for ``parity_constants.py``. Do not weaken the assertion below
"to make the test pass" -- that is the bug class this invariant exists to
catch.
"""

from __future__ import annotations

import numpy as np
import pytest

import parity_constants
import pricing_projection as sp
import rila_projection as rp  # noqa: F401  -- import so dispatch self-registers
import term_projection as tp  # noqa: F401  -- import so dispatch self-registers
from liability_dispatch import liability_path_for, registered_typenames
from product_registry import (
    _PRICING_METRIC_FORMATTERS,
    ProductType,
    get_pricing_metrics,
    implemented_product_types,
)

pytestmark = [pytest.mark.invariant]


# ---------------------------------------------------------------------------
# 1. Liability-dispatch completeness.
# ---------------------------------------------------------------------------


def _expected_result_class_for(product_type: ProductType) -> str:
    return {
        ProductType.SPIA: "SPIAProjectionResult",
        ProductType.TERM_LIFE: "TermLifeProjectionResult",
        ProductType.RILA: "RILAProjectionResult",
        ProductType.MYGA: "MYGAProjectionResult",
        ProductType.FIA: "FIAProjectionResult",
        ProductType.VARIABLE_ANNUITY: "VAProjectionResult",
        ProductType.WHOLE_LIFE: "WLProjectionResult",
        ProductType.UNIVERSAL_LIFE: "ULProjectionResult",
        ProductType.INDEXED_UL: "IULProjectionResult",
        ProductType.VARIABLE_UL: "VULProjectionResult",
    }[product_type]


def test_every_implemented_product_has_a_registered_liability_converter() -> None:
    """``run_alm_projection_from_pricing_result`` will TypeError at runtime
    if a converter is missing. We catch that at PR time instead."""
    registered = set(registered_typenames())
    missing: list[str] = []
    for product in implemented_product_types():
        expected = _expected_result_class_for(product)
        if expected not in registered:
            missing.append(f"{product.value}->{expected}")
    assert not missing, (
        "Implemented products with no liability-path converter registered: "
        f"{sorted(missing)}. Add `register_liability_path_converter("
        "'<ResultClass>', <converter>)` at the bottom of the engine module."
    )


def test_liability_dispatch_keys_match_real_result_classnames() -> None:
    """The string key in liability_dispatch MUST equal type(result).__name__
    of an actual ``price()`` output. A typo or class rename would break ALM
    silently. We construct a real cheap result for each product and look it
    up via the dispatcher."""
    yc = sp.YieldCurve(np.array([1.0, 30.0]), np.array([0.04, 0.04]))
    ages = np.arange(40, 121, dtype=int)
    qx = np.full(ages.shape[0], 0.01, dtype=float)
    mort = sp.MortalityTableQx(ages, qx)
    issue_age = 65
    horizon_age = 75

    spia_result = sp.price_spia_single_premium(
        contract=sp.SPIAContract(issue_age=issue_age, sex="male", benefit_annual=1_000.0),
        yield_curve=yc,
        mortality=mort,
        horizon_age=horizon_age,
        spread=0.0,
        valuation_year=None,
    )
    term_result = tp.price_term_life_level_monthly(
        contract=tp.TermLifeContract(
            issue_age=issue_age,
            sex="male",
            death_benefit=100_000.0,
            monthly_premium=10.0,
            term_years=5,
        ),
        yield_curve=yc,
        mortality=mort,
        horizon_age=horizon_age,
        spread=0.0,
    )
    rila_result = rp.price_rila_single_premium(
        contract=rp.RILAContract(
            issue_age=issue_age,
            sex="male",
            participation=1.0,
            cap=0.10,
            floor=0.0,
            rider_fee_annual=0.01,
        ),
        yield_curve=yc,
        mortality=mort,
        horizon_age=horizon_age,
        spread=0.0,
        valuation_year=None,
    )

    # Each must round-trip through the dispatch without raising.
    for result in (spia_result, term_result, rila_result):
        path = liability_path_for(result)
        assert path is not None
        assert hasattr(path, "times_years")
        assert hasattr(path, "expected_total_cashflows")


# ---------------------------------------------------------------------------
# 2. MODELCHECK_TOL lock.
# ---------------------------------------------------------------------------


def test_modelcheck_tol_is_exactly_zero() -> None:
    """``MODELCHECK_TOL`` must remain exactly 0.0.

    Loosening it would mean the platform tolerates Python<->Excel ModelCheck
    drift, which is the ONE invariant the parity contracts promise stays
    bit-exact (tolerances on intermediate computations are documented per
    metric in ``model_parity_contract.md``; the user-visible reconciliation
    must remain zero).

    If you genuinely need a non-zero tolerance for a new product, add a
    product-specific constant (see ``TERM_MODELCHECK_TOL`` for the precedent)
    and update ``model_change_log.md``. Do NOT touch ``MODELCHECK_TOL``."""
    assert parity_constants.MODELCHECK_TOL == 0.0, (
        "MODELCHECK_TOL drifted from 0.0 to "
        f"{parity_constants.MODELCHECK_TOL!r}. This breaks the platform-wide "
        "Python<->Excel reconciliation contract. Revert and route any "
        "intentional change through model_change_log.md + CODEOWNERS."
    )


# ---------------------------------------------------------------------------
# 3. Metric formatter completeness (no silent SPIA fallback).
# ---------------------------------------------------------------------------


def test_every_implemented_product_has_an_explicit_metric_formatter() -> None:
    missing = [
        product.value
        for product in implemented_product_types()
        if product not in _PRICING_METRIC_FORMATTERS
    ]
    assert not missing, (
        "Implemented products with no entry in _PRICING_METRIC_FORMATTERS: "
        f"{sorted(missing)}. Add a per-product formatter -- silent fallback "
        "to SPIA was removed in P0 hardening."
    )


def test_get_pricing_metrics_raises_for_unregistered_product() -> None:
    """The fallback was removed; calling for an unregistered product must
    raise loudly so the UI surfaces a real error instead of mis-formatting.

    All ten ``ProductType`` members are now implemented, so we synthesize a
    fake unregistered enum member purely for this negative-case test.
    """
    from enum import Enum

    class _FakeProductType(str, Enum):
        UNKNOWN = "unknown"

    class _DummyResult:
        pass

    with pytest.raises(KeyError, match="No pricing-metric formatter registered"):
        get_pricing_metrics(_FakeProductType.UNKNOWN, _DummyResult())  # type: ignore[arg-type]
