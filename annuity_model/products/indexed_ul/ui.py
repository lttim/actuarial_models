"""INDEXED_UL adapter, metric formatter, and UI config (re-export shim)."""

from __future__ import annotations

from product_registry import (
    ProductType,
    _PRICING_METRIC_FORMATTERS,
    get_product_adapter,
    get_product_ui_config,
)

INDEXED_UL_ADAPTER = get_product_adapter(ProductType.INDEXED_UL)
indexed_ul_metric_formatter = _PRICING_METRIC_FORMATTERS[ProductType.INDEXED_UL]
indexed_ul_ui_config = get_product_ui_config(ProductType.INDEXED_UL)


__all__ = [
    "INDEXED_UL_ADAPTER",
    "indexed_ul_metric_formatter",
    "indexed_ul_ui_config",
]
