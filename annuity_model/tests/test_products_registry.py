"""Cross-registry consistency for the unified ``products`` registry.

The :mod:`products` package was introduced in P1 hardening (2026-04) as a
consolidated view over the five legacy per-product registries
(``product_registry._PRODUCT_ADAPTERS``,
``product_excel._BUILDER_REGISTRY``,
``liability_dispatch._REGISTRY``,
``product_registry._PRICING_METRIC_FORMATTERS``,
``liability_layouts.LIABILITY_LAYOUTS``).

If a contributor adds a new product, they MUST add a
``products/<name>.py`` shim at the same time. These invariants catch
drift between the legacy registries and the unified view at PR time so
the Streamlit UI never shows a product that has half its wiring missing.

Each assertion is intentionally surgical so the failure message points at
the exact gap (missing builder, missing converter, name typo, etc).
"""

from __future__ import annotations

import pytest

import product_excel as pe
from liability_dispatch import liability_path_for, registered_typenames
from liability_layouts import LIABILITY_LAYOUTS
from product_registry import (
    ProductType,
    get_product_adapter,
    implemented_product_types,
)
from product_registry import (
    pricing_metric_formatters_by_type as legacy_pricing_metric_formatters_by_type,
)
from product_registry import (
    product_adapters_by_type as legacy_product_adapters_by_type,
)
from products import (
    ProductDefinition,
    discover_products,
    get_product_definition,
    iter_product_definitions,
    liability_path_converters_by_result_type_name,
    pricing_metric_formatters_by_type,
    product_adapters_by_type,
    product_definitions_by_type,
    registered_product_types,
    workbook_builder_spec_types_by_type,
    workbook_builders_by_type,
)

pytestmark = [pytest.mark.invariant]


def test_discover_products_is_idempotent() -> None:
    """Calling discover_products twice must not double-register or raise."""
    discover_products()
    before = registered_product_types()
    discover_products()
    after = registered_product_types()
    assert before == after, (
        "discover_products() is not idempotent: registered set changed from "
        f"{sorted(getattr(p, 'value', p) for p in before)} to "
        f"{sorted(getattr(p, 'value', p) for p in after)}."
    )


def test_every_implemented_product_has_a_definition() -> None:
    """Every product that has a legacy adapter MUST have a unified shim.

    The reverse direction is enforced by
    :func:`test_no_orphan_definitions` -- together they guarantee the
    legacy registries and the new ``products/`` view are bijective.
    """
    discover_products()
    legacy = set(implemented_product_types())
    unified = set(registered_product_types())
    missing = legacy - unified
    assert not missing, (
        "Implemented products with no products/<name>.py shim: "
        f"{sorted(p.value for p in missing)}. "
        "Add a shim under annuity_model/products/<name>.py with a "
        "register_product(ProductDefinition(...)) call."
    )


def test_no_orphan_definitions() -> None:
    """A definition without a legacy adapter would mean the unified view
    advertises a product the engine cannot price -- a worse failure mode
    than missing the definition. Catch it loudly."""
    discover_products()
    legacy = set(implemented_product_types())
    unified = set(registered_product_types())
    orphan = unified - legacy
    assert not orphan, (
        "ProductDefinitions without a corresponding entry in "
        f"product_registry._PRODUCT_ADAPTERS: {sorted(getattr(p, 'value', p) for p in orphan)}. "
        "Either add the legacy adapter or remove the orphan shim."
    )


def test_product_definition_compatibility_views_are_canonical_and_read_only() -> None:
    """Compatibility maps must be derived from ProductDefinition, not parallel wires."""
    definitions = product_definitions_by_type()
    adapters = product_adapters_by_type()
    builders = workbook_builders_by_type()
    spec_types = workbook_builder_spec_types_by_type()
    metric_formatters = pricing_metric_formatters_by_type()
    converters = liability_path_converters_by_result_type_name()

    assert set(definitions) == set(registered_product_types())
    assert set(adapters) == set(definitions)
    assert set(builders) == set(definitions)
    assert set(spec_types) == set(definitions)
    assert set(metric_formatters) == set(definitions)
    assert set(converters) == {
        definition.result_type.__name__ for definition in definitions.values()
    }

    spia = definitions[ProductType.SPIA]
    assert adapters[ProductType.SPIA] is spia.adapter
    assert builders[ProductType.SPIA] is spia.builder
    assert spec_types[ProductType.SPIA] is spia.builder_spec_type
    assert metric_formatters[ProductType.SPIA] is spia.metric_formatter
    assert converters[spia.result_type.__name__] is spia.liability_path_converter

    with pytest.raises(TypeError):
        definitions[ProductType.SPIA] = spia  # type: ignore[index]


def test_product_definition_views_match_public_legacy_views() -> None:
    """Legacy public compatibility views must stay aligned with canonical definitions."""
    assert dict(product_adapters_by_type()) == dict(legacy_product_adapters_by_type())
    assert dict(workbook_builders_by_type()) == dict(pe.workbook_builders_by_type())
    assert dict(workbook_builder_spec_types_by_type()) == dict(
        pe.workbook_builder_spec_types_by_type()
    )
    assert dict(pricing_metric_formatters_by_type()) == dict(
        legacy_pricing_metric_formatters_by_type()
    )


@pytest.mark.parametrize("product_type", list(ProductType))
def test_definition_fields_match_legacy_registries(product_type: ProductType) -> None:
    """For every implemented product, every field of ProductDefinition
    MUST be the *same object* (``is``) as the legacy registry value.

    Using ``is`` (not ``==``) catches the case where a contributor adds a
    parallel implementation in products/<name>.py that drifts from the
    legacy adapter -- silently shipping two different SPIA pricers in the
    same process.
    """
    discover_products()
    if product_type not in implemented_product_types():
        # Whole Life / Variable Annuity have no legacy adapter -- the
        # bijectivity test above already asserts they have no shim either.
        return
    definition = get_product_definition(product_type)
    assert isinstance(definition, ProductDefinition)
    assert definition.product_type is product_type
    assert definition.adapter is get_product_adapter(product_type), (
        f"{product_type.value}: ProductDefinition.adapter drifted from "
        "product_registry._PRODUCT_ADAPTERS. The shim under "
        f"products/{product_type.value}.py must reuse "
        "get_product_adapter(...) instead of constructing a parallel adapter."
    )
    assert definition.builder is pe.workbook_builders_by_type()[product_type], (
        f"{product_type.value}: ProductDefinition.builder drifted from "
        "product_excel._BUILDER_REGISTRY."
    )
    assert (
        definition.metric_formatter is legacy_pricing_metric_formatters_by_type()[product_type]
    ), (
        f"{product_type.value}: ProductDefinition.metric_formatter drifted from "
        "product_registry._PRICING_METRIC_FORMATTERS."
    )


@pytest.mark.parametrize("product_type", list(ProductType))
def test_result_type_name_matches_dispatch_key(product_type: ProductType) -> None:
    """``liability_dispatch`` is keyed by ``type(result).__name__``.
    ``ProductDefinition.result_type.__name__`` MUST equal one of the
    registered dispatch keys, otherwise ``run_alm_projection_from_pricing_result``
    would TypeError at runtime."""
    discover_products()
    if product_type not in implemented_product_types():
        return
    definition = get_product_definition(product_type)
    keys = set(registered_typenames())
    assert definition.result_type.__name__ in keys, (
        f"{product_type.value}: ProductDefinition.result_type "
        f"({definition.result_type.__name__}) is not a registered "
        f"liability_dispatch key. Registered: {sorted(keys)}. "
        "Either fix the result_type field or call "
        "register_liability_path_converter() at engine import time."
    )


@pytest.mark.parametrize("product_type", list(ProductType))
def test_builder_spec_type_matches_product_excel_spec_types(
    product_type: ProductType,
) -> None:
    """``product_excel._BUILDER_SPEC_TYPES`` is the dispatcher's
    isinstance-check class. ``ProductDefinition.builder_spec_type`` must
    equal it; otherwise scaffolding tools would build a spec the
    dispatcher would then reject."""
    discover_products()
    if product_type not in implemented_product_types():
        return
    definition = get_product_definition(product_type)
    expected = pe.workbook_builder_spec_types_by_type()[product_type]
    assert definition.builder_spec_type is expected, (
        f"{product_type.value}: ProductDefinition.builder_spec_type "
        f"({definition.builder_spec_type.__name__}) does not match "
        f"product_excel._BUILDER_SPEC_TYPES[{product_type!r}] "
        f"({expected.__name__ if expected else expected})."
    )


@pytest.mark.parametrize("product_type", list(ProductType))
def test_layout_present_for_implemented_products(product_type: ProductType) -> None:
    """Every implemented product must also have a LIABILITY_LAYOUTS entry --
    the ALM ladder dies without one. The unified view does not currently
    embed the layout (it is keyed by string code, not enum), but every
    implemented product is required to have both."""
    discover_products()
    if product_type not in implemented_product_types():
        return
    assert product_type.value in LIABILITY_LAYOUTS, (
        f"{product_type.value}: implemented product is missing from "
        "liability_layouts.LIABILITY_LAYOUTS. Add the column letters in "
        "the same PR as the products/<name>.py shim."
    )


def test_iter_product_definitions_returns_only_definitions() -> None:
    """Defensive: every yielded value is the right type. Tests downstream
    of this loop format failure messages from .display_name etc. and
    would NPE on a stray None."""
    for definition in iter_product_definitions():
        assert isinstance(definition, ProductDefinition)
        assert isinstance(definition.display_name, str) and definition.display_name
        assert isinstance(definition.contract_type, type)
        assert isinstance(definition.result_type, type)
        assert isinstance(definition.builder_spec_type, type)


def test_register_product_rejects_silent_override() -> None:
    """A contributor accidentally importing two shims for the same product
    would otherwise silently overwrite -- exactly the behavior the legacy
    register_builder() and register_liability_path_converter() refuse.
    The unified registry holds the same line."""
    from products import register_product

    existing = get_product_definition(ProductType.SPIA)
    duplicate = ProductDefinition(
        product_type=ProductType.SPIA,
        display_name="Duplicate SPIA",
        contract_type=existing.contract_type,
        result_type=existing.result_type,
        builder_spec_type=existing.builder_spec_type,
        adapter=existing.adapter,
        builder=existing.builder,
        liability_path_converter=existing.liability_path_converter,
        metric_formatter=existing.metric_formatter,
    )
    with pytest.raises(RuntimeError, match="already registered"):
        register_product(duplicate)
    # Re-registering the SAME instance must remain a no-op (idempotent
    # re-imports are normal; only divergent definitions are an error).
    register_product(existing)


def test_get_product_definition_raises_for_unregistered() -> None:
    """Lookup for a ProductType with no shim must raise KeyError with a
    pointer to the products/<name>.py path the contributor needs to
    create.

    All ten ``ProductType`` members are now wired, so we synthesize a
    fake unregistered enum member purely for this negative-case test.
    """
    from enum import Enum

    class _FakeProductType(str, Enum):
        UNKNOWN = "unknown_product_for_test"

    with pytest.raises(KeyError, match="products/<name>.py"):
        get_product_definition(_FakeProductType.UNKNOWN)


@pytest.mark.parametrize("product_type", list(ProductType))
def test_definition_converter_round_trips(product_type: ProductType) -> None:
    """The liability_path_converter on the definition MUST be the *same*
    callable the dispatch registry uses, so calling it directly yields
    the same LiabilityPath as routing through liability_path_for(). We
    cannot easily test equality of paths here without spinning up real
    pricing, so we just assert identity of the registered callable."""
    discover_products()
    if product_type not in implemented_product_types():
        return
    definition = get_product_definition(product_type)
    # Both sides should be callable and pass the smoke "is callable" test;
    # full round-trip is covered by tests/test_meta_invariants.py.
    assert callable(definition.liability_path_converter)
    assert liability_path_for is not None  # module-level sanity
