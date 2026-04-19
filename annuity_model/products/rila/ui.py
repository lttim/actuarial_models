"""RILA adapter, metric formatter, and UI config (re-export shim)."""

from __future__ import annotations

from product_registry import (
    ProductType,
    _PRICING_METRIC_FORMATTERS,
    get_product_adapter,
    get_product_ui_config,
)

RILA_ADAPTER = get_product_adapter(ProductType.RILA)
rila_metric_formatter = _PRICING_METRIC_FORMATTERS[ProductType.RILA]
rila_ui_config = get_product_ui_config(ProductType.RILA)


__all__ = [
    "RILA_ADAPTER",
    "rila_metric_formatter",
    "rila_ui_config",
]
