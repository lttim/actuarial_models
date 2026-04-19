"""VARIABLE_UL adapter, metric formatter, and UI config (re-export shim)."""

from __future__ import annotations

from product_registry import (
    ProductType,
    _PRICING_METRIC_FORMATTERS,
    get_product_adapter,
    get_product_ui_config,
)

VARIABLE_UL_ADAPTER = get_product_adapter(ProductType.VARIABLE_UL)
variable_ul_metric_formatter = _PRICING_METRIC_FORMATTERS[ProductType.VARIABLE_UL]
variable_ul_ui_config = get_product_ui_config(ProductType.VARIABLE_UL)


__all__ = [
    "VARIABLE_UL_ADAPTER",
    "variable_ul_metric_formatter",
    "variable_ul_ui_config",
]
