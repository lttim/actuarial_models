"""WHOLE_LIFE adapter, metric formatter, and UI config (re-export shim)."""

from __future__ import annotations

from product_registry import (
    _PRICING_METRIC_FORMATTERS,
    ProductType,
    get_product_adapter,
    get_product_ui_config,
)

WHOLE_LIFE_ADAPTER = get_product_adapter(ProductType.WHOLE_LIFE)
whole_life_metric_formatter = _PRICING_METRIC_FORMATTERS[ProductType.WHOLE_LIFE]
whole_life_ui_config = get_product_ui_config(ProductType.WHOLE_LIFE)


__all__ = [
    "WHOLE_LIFE_ADAPTER",
    "whole_life_metric_formatter",
    "whole_life_ui_config",
]
