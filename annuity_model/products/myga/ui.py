"""MYGA adapter, metric formatter, and UI config (re-export shim)."""

from __future__ import annotations

from product_registry import (
    ProductType,
    _PRICING_METRIC_FORMATTERS,
    get_product_adapter,
    get_product_ui_config,
)

MYGA_ADAPTER = get_product_adapter(ProductType.MYGA)
myga_metric_formatter = _PRICING_METRIC_FORMATTERS[ProductType.MYGA]
myga_ui_config = get_product_ui_config(ProductType.MYGA)


__all__ = [
    "MYGA_ADAPTER",
    "myga_metric_formatter",
    "myga_ui_config",
]
