"""SPIA adapter, metric formatter, and UI config (re-export shim).

Canonical import path for SPIA's UI-facing surface. The actual pricing
adapter, metric formatter, and ``ProductUIConfig`` instance live in
:mod:`product_registry`; this shim is the narrow access path that keeps
the consolidated :mod:`products.spia` package self-contained.

Notes
-----
The SPIA adapter, formatter, and UI config are read directly from
:mod:`product_registry` rather than re-implemented here. Any future
move of the adapter into this subpackage just changes the imports
below; no caller needs to update.
"""

from __future__ import annotations

from product_registry import (
    _PRICING_METRIC_FORMATTERS,
    ProductType,
    get_product_adapter,
    get_product_ui_config,
)

# Resolve at import time so the module-level constants are stable refs
# (downstream `is`-comparisons in tests rely on this).
SPIA_ADAPTER = get_product_adapter(ProductType.SPIA)
spia_metric_formatter = _PRICING_METRIC_FORMATTERS[ProductType.SPIA]
spia_ui_config = get_product_ui_config(ProductType.SPIA)


__all__ = [
    "SPIA_ADAPTER",
    "spia_metric_formatter",
    "spia_ui_config",
]
