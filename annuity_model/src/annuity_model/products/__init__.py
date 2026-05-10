"""Canonical per-product registry.

``ProductDefinition`` is the source of truth for the product-facing
platform wires: pricing adapter, Excel builder/spec, pricing-result to
liability-path converter, UI metric formatter, UI metadata, capability
flags, mortality defaults, validator hook, and user-facing order.

The older module-level registries in :mod:`product_registry`,
:mod:`product_excel`, and :mod:`liability_dispatch` remain as seed
registries while the implementation is being migrated, but public
compatibility views are derived from this module. New products should add
one ``products/<name>/__init__.py`` shim with a complete
``ProductDefinition`` and should not add new public dispatch tables.

This module exposes:

* ``register_product(definition)`` -- records the definition in the
  process-wide registry. Modules under :mod:`products` call this at
  import time. Re-registering a product raises ``RuntimeError`` -- the
  legacy registries already use the same "no silent overrides" rule, so
  this matches platform-wide behavior.

* ``iter_product_definitions()`` -- returns every registered definition.
  Triggers :func:`discover_products` lazily on first call so contributors
  do not need to manually import each ``products.<name>`` submodule.

* ``get_product_definition(product_type)`` -- single-product lookup,
  raises ``KeyError`` if the definition is missing (typically because the
  product's ``products/<name>.py`` shim has not been written yet).

* ``discover_products()`` -- imports every submodule under :mod:`products`
  via ``pkgutil.iter_modules``. Idempotent; safe to call repeatedly.

A new product authored against this module looks like::

    # products/fia.py
    from annuity_model.products import register_product, ProductDefinition
    from annuity_model.product_registry import ProductType, ProductCapabilities, ...
    from annuity_model.product_excel import _BUILDER_REGISTRY
    ...

    register_product(
        ProductDefinition(
            product_type=ProductType.FIA,
            display_name="FIA (S&P-linked, deterministic)",
            ...,
        )
    )

The accompanying ``tests/test_products_registry.py`` enforces that the
canonical view feeds the legacy public compatibility functions.
"""

from __future__ import annotations

import importlib
import pkgutil
import threading
from collections.abc import Callable, Iterator, Mapping
from dataclasses import dataclass
from types import MappingProxyType
from typing import Any

# Type aliases kept very loose at this layer because the source registries
# (product_registry, product_excel, liability_dispatch) intentionally type
# their callables as ``Callable[..., Any]`` to avoid a forest of generics.
# The per-product engine/builder retain their narrow types; this façade
# only needs to prove "something is wired" for the meta-tests.
PricingMetricFormatter = Callable[[Any], Any]
LiabilityPathConverter = Callable[[Any], Any]
WorkbookBuilder = Callable[..., bytes]
RunInputValidator = Callable[[Mapping[str, Any]], list[str]]


@dataclass(frozen=True, slots=True)
class ProductDefinition:
    """All per-product wires bundled into one immutable record.

    Attributes
    ----------
    product_type:
        The :class:`product_registry.ProductType` enum member that keys this
        definition. Treated as the canonical identity by the registry.
    display_name:
        Human-readable label, e.g. "SPIA". Mirrors
        :func:`product_registry.product_label`.
    contract_type:
        The pricing-input dataclass class (e.g. ``SPIAContract``). Used by
        scaffolding tooling and by AppTest UI smoke tests to construct
        contracts without going through the registry.
    result_type:
        The pricing-output class (e.g. ``SPIAProjectionResult``). The
        liability_dispatch key MUST equal ``result_type.__name__``;
        :func:`tests.test_products_registry` enforces this.
    builder_spec_type:
        The Excel build-spec dataclass (e.g. ``ExcelBuildSpec``,
        ``TermExcelBuildSpec``, ``RILAExcelBuildSpec``). The dispatcher in
        :func:`product_excel.build_product_workbook` does an isinstance
        check against this class.
    adapter:
        The :class:`product_registry.ProductAdapter` instance --
        ``price()``, ``price_monte_carlo()``, ``excel_spec_from_run()``.
    builder:
        The Excel workbook builder callable registered in
        :data:`product_excel._BUILDER_REGISTRY`.
    liability_path_converter:
        Callable from pricing result to ``LiabilityPath``. Same callable
        registered in :mod:`liability_dispatch`.
    metric_formatter:
        Callable from pricing result to the UI metric tuple. Public metric
        formatter lookup is derived from this field.
    capabilities:
        Product capability flags used by UI/CLI surfaces.
    ui_config:
        Product-specific UI copy and download filenames.
    mortality_mode_options:
        Supported mortality mode keys for the product.
    default_mortality_mode:
        Default mortality mode key selected by the UI.
    validator:
        Optional run-form validator. ``None`` means the product only uses
        the cross-product launch checks.
    order:
        Stable UI/CLI product ordering. This replaces hard-coded option
        lists in the legacy registry.
    maturity_label:
        Short status string used by UX surfaces and future release notes.
    assumption_profile:
        Assumption evidence status for the product. Current values are
        descriptive strings, not enforcement gates; Chunk 6 will persist
        assumption evidence in the run ledger.

    Notes
    -----
    This dataclass is **frozen + slotted** because it is consumed by tests
    that walk every product and format failure messages from its fields;
    accidentally rebinding an attribute mid-test would mask the kind of
    drift these invariants are meant to catch.
    """

    product_type: Any  # ProductType, kept Any to avoid circular import
    display_name: str
    contract_type: type
    result_type: type
    builder_spec_type: type
    adapter: Any
    builder: WorkbookBuilder
    liability_path_converter: LiabilityPathConverter
    metric_formatter: PricingMetricFormatter
    capabilities: Any
    ui_config: Any
    mortality_mode_options: tuple[str, ...]
    default_mortality_mode: str
    validator: RunInputValidator | None
    order: int
    maturity_label: str = "Mechanics-production"
    assumption_profile: str = "demo-safe-with-waiver"


_DEFINITIONS: dict[Any, ProductDefinition] = {}
_DISCOVERY_LOCK = threading.Lock()
_DISCOVERED = False


def register_product(definition: ProductDefinition) -> ProductDefinition:
    """Record *definition* under ``definition.product_type``.

    Returns the definition unchanged so the call site can also bind it to
    a module-level constant if it wishes (``DEF = register_product(...)``).

    Raises
    ------
    RuntimeError
        If a definition is already registered under the same product_type
        AND it is not the same dataclass instance. This matches the
        no-silent-override discipline of the legacy registries
        (``product_excel.register_builder``,
        ``liability_dispatch.register_liability_path_converter``).
    """
    existing = _DEFINITIONS.get(definition.product_type)
    if existing is not None and existing is not definition:
        raise RuntimeError(
            f"ProductDefinition for {definition.product_type!r} is already "
            f"registered (existing display_name={existing.display_name!r}, "
            f"new display_name={definition.display_name!r}). "
            "Re-importing the same products.<name> module is fine, but "
            "two different definitions for the same ProductType is almost "
            "always a copy-paste bug."
        )
    _DEFINITIONS[definition.product_type] = definition
    return definition


def discover_products() -> None:
    """Import every submodule under :mod:`products` so it can self-register.

    Idempotent and thread-safe. Submodules whose import fails are NOT
    silently swallowed -- the exception propagates so a broken product
    shim surfaces as a hard error at the first :func:`iter_product_definitions`
    call instead of silently dropping that product from the UI.
    """
    global _DISCOVERED
    if _DISCOVERED:
        return
    with _DISCOVERY_LOCK:
        if _DISCOVERED:
            return
        package = importlib.import_module(__name__)
        # __path__ is set on packages; ignore on type-checkers that don't
        # know that.
        for _finder, name, _is_pkg in pkgutil.iter_modules(package.__path__):
            if name.startswith("_"):
                continue
            importlib.import_module(f"{__name__}.{name}")
        _DISCOVERED = True


def iter_product_definitions() -> Iterator[ProductDefinition]:
    """Yield every registered :class:`ProductDefinition` (auto-discovers).

    Order is registration order, which under :func:`discover_products` is
    the alphabetical order returned by :func:`pkgutil.iter_modules`. Tests
    sort by ``product_type.value`` when they need a stable order.
    """
    discover_products()
    yield from _DEFINITIONS.values()


def get_product_definition(product_type: Any) -> ProductDefinition:
    """Return the registered definition for *product_type* (auto-discovers).

    Raises
    ------
    KeyError
        If no products/<name>.py registers this product. Most likely the
        contributor added a new ``ProductType`` member but forgot the shim
        under :mod:`products`. Add ``products/<name>.py`` with a
        ``register_product(ProductDefinition(...))`` call.
    """
    discover_products()
    try:
        return _DEFINITIONS[product_type]
    except KeyError as exc:
        registered = sorted(getattr(p, "value", repr(p)) for p in _DEFINITIONS)
        raise KeyError(
            f"No ProductDefinition registered for {product_type!r}. "
            f"Registered products: {registered}. Add a shim under "
            f"`annuity_model/products/<name>.py` that calls "
            "`register_product(ProductDefinition(...))`."
        ) from exc


def registered_product_types() -> tuple[Any, ...]:
    """Return the tuple of ProductType keys (auto-discovers).

    Used by tests/test_products_registry to walk every registered product
    in deterministic order.
    """
    discover_products()
    return tuple(_DEFINITIONS)


def product_definitions_by_type() -> Mapping[Any, ProductDefinition]:
    """Read-only canonical ProductDefinition view keyed by product type.

    This is the compatibility seam for legacy registry consumers during the
    migration to ``ProductDefinition`` as the source of truth. Callers that
    only need a product-indexed view should prefer this immutable mapping
    over reaching into the private legacy dictionaries.
    """
    discover_products()
    return MappingProxyType(dict(_DEFINITIONS))


def product_adapters_by_type() -> Mapping[Any, Any]:
    """Read-only adapter compatibility view derived from definitions."""
    return MappingProxyType(
        {
            product_type: definition.adapter
            for product_type, definition in product_definitions_by_type().items()
        }
    )


def product_display_names_by_type() -> Mapping[Any, str]:
    """Read-only display-name compatibility view derived from definitions."""
    return MappingProxyType(
        {
            product_type: definition.display_name
            for product_type, definition in product_definitions_by_type().items()
        }
    )


def product_capabilities_by_type() -> Mapping[Any, Any]:
    """Read-only capability compatibility view derived from definitions."""
    return MappingProxyType(
        {
            product_type: definition.capabilities
            for product_type, definition in product_definitions_by_type().items()
        }
    )


def product_ui_configs_by_type() -> Mapping[Any, Any]:
    """Read-only UI-config compatibility view derived from definitions."""
    return MappingProxyType(
        {
            product_type: definition.ui_config
            for product_type, definition in product_definitions_by_type().items()
        }
    )


def product_mortality_mode_options_by_type() -> Mapping[Any, tuple[str, ...]]:
    """Read-only mortality-options view derived from definitions."""
    return MappingProxyType(
        {
            product_type: definition.mortality_mode_options
            for product_type, definition in product_definitions_by_type().items()
        }
    )


def product_default_mortality_modes_by_type() -> Mapping[Any, str]:
    """Read-only mortality-default view derived from definitions."""
    return MappingProxyType(
        {
            product_type: definition.default_mortality_mode
            for product_type, definition in product_definitions_by_type().items()
        }
    )


def product_validators_by_type() -> Mapping[Any, RunInputValidator]:
    """Read-only per-product validator view derived from definitions."""
    return MappingProxyType(
        {
            product_type: definition.validator
            for product_type, definition in product_definitions_by_type().items()
            if definition.validator is not None
        }
    )


def product_options_for_ui() -> tuple[Any, ...]:
    """Stable UI product ordering derived from definitions."""
    return tuple(
        definition.product_type
        for definition in sorted(iter_product_definitions(), key=lambda item: item.order)
    )


def product_maturity_labels_by_type() -> Mapping[Any, str]:
    """Read-only product maturity status view derived from definitions."""
    return MappingProxyType(
        {
            product_type: definition.maturity_label
            for product_type, definition in product_definitions_by_type().items()
        }
    )


def product_assumption_profiles_by_type() -> Mapping[Any, str]:
    """Read-only assumption evidence status view derived from definitions."""
    return MappingProxyType(
        {
            product_type: definition.assumption_profile
            for product_type, definition in product_definitions_by_type().items()
        }
    )


def workbook_builders_by_type() -> Mapping[Any, WorkbookBuilder]:
    """Read-only workbook-builder compatibility view derived from definitions."""
    return MappingProxyType(
        {
            product_type: definition.builder
            for product_type, definition in product_definitions_by_type().items()
        }
    )


def workbook_builder_spec_types_by_type() -> Mapping[Any, type]:
    """Read-only workbook spec-type compatibility view derived from definitions."""
    return MappingProxyType(
        {
            product_type: definition.builder_spec_type
            for product_type, definition in product_definitions_by_type().items()
        }
    )


def pricing_metric_formatters_by_type() -> Mapping[Any, PricingMetricFormatter]:
    """Read-only pricing-metric compatibility view derived from definitions."""
    return MappingProxyType(
        {
            product_type: definition.metric_formatter
            for product_type, definition in product_definitions_by_type().items()
        }
    )


def liability_path_converters_by_result_type_name() -> Mapping[str, LiabilityPathConverter]:
    """Read-only liability converter view keyed like ``liability_dispatch``.

    ``liability_dispatch`` routes on ``type(result).__name__``. Deriving this
    mapping from ``ProductDefinition`` lets tests and migration code compare
    the canonical view to the legacy dispatch registry without importing the
    private dispatch dictionary.
    """
    return MappingProxyType(
        {
            definition.result_type.__name__: definition.liability_path_converter
            for definition in iter_product_definitions()
        }
    )


__all__ = [
    "LiabilityPathConverter",
    "PricingMetricFormatter",
    "ProductDefinition",
    "RunInputValidator",
    "WorkbookBuilder",
    "discover_products",
    "get_product_definition",
    "iter_product_definitions",
    "liability_path_converters_by_result_type_name",
    "pricing_metric_formatters_by_type",
    "product_adapters_by_type",
    "product_assumption_profiles_by_type",
    "product_capabilities_by_type",
    "product_default_mortality_modes_by_type",
    "product_definitions_by_type",
    "product_display_names_by_type",
    "product_maturity_labels_by_type",
    "product_mortality_mode_options_by_type",
    "product_options_for_ui",
    "product_ui_configs_by_type",
    "product_validators_by_type",
    "register_product",
    "registered_product_types",
    "workbook_builder_spec_types_by_type",
    "workbook_builders_by_type",
]
