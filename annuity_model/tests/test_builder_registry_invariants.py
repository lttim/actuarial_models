"""Workbook-builder registry invariants.

The Phase-5 hardening pass replaced the if/elif chain inside
``build_product_workbook`` with a ``@register_builder`` decorator pattern in
``product_excel.py``. The dispatcher now looks up the per-product builder by
``ProductType`` enum, so adding a new product means writing a thin builder
function and decorating it -- no editing the dispatcher itself.

These invariants lock that pattern in place:

  1. **No drift between the adapter registry and the builder registry.** Every
     product with an implemented :class:`ProductAdapter`
     (``product_registry.implemented_product_types()``) MUST also have a
     registered workbook builder. The reverse is also enforced: no orphan
     builders for products that don't have an adapter (which would mean
     dead code or a typo'd enum).
  2. **Spec-type validation fails fast.** Passing a wrong spec type to
     ``build_product_workbook`` raises ``TypeError`` with a clear message
     -- not a confusing ``AttributeError`` from inside the builder.
  3. **Re-registration is rejected.** Re-decorating an already-registered
     ``ProductType`` raises ``RuntimeError`` (catches copy-paste mistakes
     where two adapters claim the same enum).

If one of these fails, the fix usually lives in ``product_excel.py`` (a
missing ``@register_builder`` decorator on a new builder), not in this
test.
"""

from __future__ import annotations

import pytest

from build_pricing_excel_workbook import ExcelBuildSpec
from build_rila_excel_workbook import RILAExcelBuildSpec
from build_term_excel_workbook import TermExcelBuildSpec
from product_excel import (
    _BUILDER_REGISTRY,
    _BUILDER_SPEC_TYPES,
    build_product_workbook,
    register_builder,
    registered_builders,
)
from product_registry import ProductType, implemented_product_types

pytestmark = [
    pytest.mark.invariant,
    pytest.mark.product_spia,
    pytest.mark.product_term,
    pytest.mark.product_rila,
]


def test_every_implemented_adapter_has_a_registered_builder() -> None:
    """No drift between product_registry and product_excel registries."""
    adapter_types = set(implemented_product_types())
    builder_types = set(registered_builders())
    missing_builders = adapter_types - builder_types
    orphan_builders = builder_types - adapter_types
    assert not missing_builders, (
        f"Implemented adapters with no @register_builder entry in product_excel.py: "
        f"{sorted(p.value for p in missing_builders)}. Add a builder function and "
        f"decorate it with @register_builder(<ProductType>, spec_type=<SpecType>)."
    )
    assert not orphan_builders, (
        f"Registered builders with no implemented adapter in product_registry.py: "
        f"{sorted(p.value for p in orphan_builders)}. Either remove the builder or "
        f"add the adapter to _PRODUCT_ADAPTERS."
    )


def test_each_builder_declares_expected_spec_type() -> None:
    """Each registered builder must declare its expected spec dataclass.

    The dispatcher uses this to fail fast on wrong-type specs.
    """
    expected = {
        ProductType.SPIA: ExcelBuildSpec,
        ProductType.TERM_LIFE: TermExcelBuildSpec,
        ProductType.RILA: RILAExcelBuildSpec,
    }
    for product_type, spec_type in expected.items():
        assert _BUILDER_SPEC_TYPES.get(product_type) is spec_type, (
            f"Builder for {product_type.value} should declare spec_type={spec_type.__name__} "
            f"but got {_BUILDER_SPEC_TYPES.get(product_type)}."
        )


def test_dispatcher_rejects_wrong_spec_type_with_clear_typeerror() -> None:
    class _NotASpec:
        pass

    with pytest.raises(TypeError, match=r"spia workbook builder requires ExcelBuildSpec"):
        build_product_workbook(product_type=ProductType.SPIA, spec=_NotASpec())
    with pytest.raises(TypeError, match=r"term_life workbook builder requires TermExcelBuildSpec"):
        build_product_workbook(product_type=ProductType.TERM_LIFE, spec=_NotASpec())
    with pytest.raises(TypeError, match=r"rila workbook builder requires RILAExcelBuildSpec"):
        build_product_workbook(product_type=ProductType.RILA, spec=_NotASpec())


def test_dispatcher_rejects_unimplemented_product_with_notimplementederror() -> None:
    class _NotASpec:
        pass

    with pytest.raises(NotImplementedError, match=r"whole_life"):
        build_product_workbook(product_type=ProductType.WHOLE_LIFE, spec=_NotASpec())
    with pytest.raises(NotImplementedError, match=r"variable_annuity"):
        build_product_workbook(product_type=ProductType.VARIABLE_ANNUITY, spec=_NotASpec())


def test_re_registering_same_product_type_raises_runtimeerror() -> None:
    """Two builders fighting for the same ProductType is almost always a bug."""

    @register_builder(ProductType.WHOLE_LIFE, spec_type=ExcelBuildSpec)
    def _placeholder_builder(**_kwargs: object) -> bytes:  # pragma: no cover -- never called
        return b""

    try:
        with pytest.raises(RuntimeError, match=r"already registered"):

            @register_builder(ProductType.WHOLE_LIFE, spec_type=ExcelBuildSpec)
            def _duplicate_builder(**_kwargs: object) -> bytes:  # pragma: no cover
                return b""

    finally:
        # Clean up so subsequent test runs / parallel workers see a pristine
        # registry. Without this, the second invocation of the test (e.g.
        # pytest --lf) would fail on the first decorator instead of the
        # intended second one.
        _BUILDER_REGISTRY.pop(ProductType.WHOLE_LIFE, None)
        _BUILDER_SPEC_TYPES.pop(ProductType.WHOLE_LIFE, None)
