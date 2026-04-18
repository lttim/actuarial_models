"""Pricing-result -> LiabilityPath dispatch registry.

Replaces the ``isinstance`` chain + deferred imports that previously lived in
:func:`pricing_projection.run_alm_projection_from_pricing_result`. Each
engine module (`pricing_projection`, `term_projection`, `rila_projection`)
calls :func:`register_liability_path_converter` at import time so the engine
core can dispatch generically without importing the per-product modules.

This is intentionally a standalone module (no project imports) so it is
safe to import from anywhere in the dependency graph without triggering
the engine <-> product cycle. New products plug in by importing this
module and registering their converter -- no edits to the core engine.
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
    converter = _REGISTRY.get(typename)
    if converter is None:
        registered = ", ".join(sorted(_REGISTRY)) or "<none>"
        raise TypeError(
            f"No liability-path converter registered for "
            f"{typename!r}; registered types: {registered}. "
            f"Engine modules must call register_liability_path_converter() "
            f"at import time."
        )
    return converter(pricing_result)


def registered_typenames() -> list[str]:
    """Return the sorted list of registered pricing-result class names.

    Used by meta-tests in Phase 4 to assert every product engine is wired in.
    """
    return sorted(_REGISTRY)


__all__ = [
    "LiabilityPathConverter",
    "liability_path_for",
    "register_liability_path_converter",
    "registered_typenames",
]
