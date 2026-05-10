"""Pricing-result -> LiabilityPath dispatch.

Replaces the ``isinstance`` chain + deferred imports that previously lived in
:func:`pricing_projection.run_alm_projection_from_pricing_result`. Engine
modules still call :func:`register_liability_path_converter` at import time as
the construction seed, but public dispatch now resolves through canonical
``ProductDefinition`` records.

The private ``_REGISTRY`` is retained for the engine registration decorator
pattern and for compatibility diagnostics. New products should expose the same
converter through their ``ProductDefinition``; public lookups below use the
definition-derived view.
"""

from __future__ import annotations

from collections.abc import Callable
from typing import Any

# Type signature: pricing_result -> LiabilityPath. Kept untyped here to
# avoid pulling in pricing_projection at module load time (LiabilityPath
# is defined there). Callers see the proper type via the engine's wrapper.
LiabilityPathConverter = Callable[[Any], Any]

_REGISTRY: dict[str, LiabilityPathConverter] = {}


def register_liability_path_converter(typename: str, converter: LiabilityPathConverter) -> None:
    """Register a converter for a pricing-result class.

    Parameters
    ----------
    typename:
        ``type(pricing_result).__name__`` (e.g. ``"SPIAProjectionResult"``).
        Using the unqualified name avoids importing the engine class here.
    converter:
        Callable that takes the pricing result and returns its
        :class:`pricing_projection.LiabilityPath`.

    Raises
    ------
    ValueError
        If a converter is already registered under the same name and is not
        the same callable. Re-registering the *same* converter (e.g. on
        module reload) is a no-op.
    """
    existing = _REGISTRY.get(typename)
    if existing is not None and existing is not converter:
        raise ValueError(
            f"Liability-path converter for {typename!r} is already registered "
            f"({existing!r}); refusing to overwrite."
        )
    _REGISTRY[typename] = converter


def liability_path_for(pricing_result: Any) -> Any:
    """Look up and invoke the converter for *pricing_result*.

    Falls back to a clear error message that lists the registered types so
    a missing registration surfaces as a TypeError instead of silently
    routing through a wrong branch.
    """
    typename = type(pricing_result).__name__
    from annuity_model.products import liability_path_converters_by_result_type_name

    converters = liability_path_converters_by_result_type_name()
    converter = converters.get(typename)
    if converter is None:
        registered = ", ".join(sorted(converters)) or "<none>"
        raise TypeError(
            f"No liability-path converter registered for "
            f"{typename!r}; registered types: {registered}. "
            f"Add the result type and converter to the product's "
            f"ProductDefinition."
        )
    return converter(pricing_result)


def registered_typenames() -> list[str]:
    """Return the sorted list of registered pricing-result class names.

    Used by meta-tests to assert every product engine is wired in.
    """
    from annuity_model.products import liability_path_converters_by_result_type_name

    return sorted(liability_path_converters_by_result_type_name())


__all__ = [
    "LiabilityPathConverter",
    "liability_path_for",
    "register_liability_path_converter",
    "registered_typenames",
]
