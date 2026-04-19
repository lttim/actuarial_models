"""UNIVERSAL_LIFE adapter, metric formatter, and UI config (re-export shim)."""

from __future__ import annotations

from product_registry import (
    ProductType,
    _PRICING_METRIC_FORMATTERS,
    get_product_adapter,
    get_product_ui_config,
)

UNIVERSAL_LIFE_ADAPTER = get_product_adapter(ProductType.UNIVERSAL_LIFE)
universal_life_metric_formatter = _PRICING_METRIC_FORMATTERS[ProductType.UNIVERSAL_LIFE]
universal_life_ui_config = get_product_ui_config(ProductType.UNIVERSAL_LIFE)


__all__ = [
    "UNIVERSAL_LIFE_ADAPTER",
    "universal_life_metric_formatter",
    "universal_life_ui_config",
]
