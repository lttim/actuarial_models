"""VARIABLE_ANNUITY adapter, metric formatter, and UI config (re-export shim)."""

from __future__ import annotations

from product_registry import (
    ProductType,
    _PRICING_METRIC_FORMATTERS,
    get_product_adapter,
    get_product_ui_config,
)

VARIABLE_ANNUITY_ADAPTER = get_product_adapter(ProductType.VARIABLE_ANNUITY)
variable_annuity_metric_formatter = _PRICING_METRIC_FORMATTERS[ProductType.VARIABLE_ANNUITY]
variable_annuity_ui_config = get_product_ui_config(ProductType.VARIABLE_ANNUITY)


__all__ = [
    "VARIABLE_ANNUITY_ADAPTER",
    "variable_annuity_metric_formatter",
    "variable_annuity_ui_config",
]
