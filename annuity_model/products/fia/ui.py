"""FIA adapter, metric formatter, and UI config (re-export shim)."""

from __future__ import annotations

from product_registry import (
    _PRICING_METRIC_FORMATTERS,
    ProductType,
    get_product_adapter,
    get_product_ui_config,
)

FIA_ADAPTER = get_product_adapter(ProductType.FIA)
fia_metric_formatter = _PRICING_METRIC_FORMATTERS[ProductType.FIA]
fia_ui_config = get_product_ui_config(ProductType.FIA)


__all__ = [
    "FIA_ADAPTER",
    "fia_metric_formatter",
    "fia_ui_config",
]
