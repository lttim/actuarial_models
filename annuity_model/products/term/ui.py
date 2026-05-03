"""Term Life adapter, metric formatter, and UI config (re-export shim).

Re-exports the Term-Life-specific UI surface (selectbox label parsers,
``TermContractUIConfig`` instance, etc.) under the canonical
``products.term.ui`` namespace.
"""

from __future__ import annotations

from product_registry import (
    _PRICING_METRIC_FORMATTERS,
    ProductType,
    get_product_adapter,
    get_product_ui_config,
    get_term_contract_ui_config,
    parse_term_benefit_timing_label,
    parse_term_length_label_to_years,
    parse_term_premium_mode_label,
    term_benefit_timing_label_options,
    term_length_label_options,
    term_premium_mode_label_options,
)

TERM_ADAPTER = get_product_adapter(ProductType.TERM_LIFE)
term_metric_formatter = _PRICING_METRIC_FORMATTERS[ProductType.TERM_LIFE]
term_ui_config = get_product_ui_config(ProductType.TERM_LIFE)
term_contract_ui_config = get_term_contract_ui_config()


__all__ = [
    "TERM_ADAPTER",
    "parse_term_benefit_timing_label",
    "parse_term_length_label_to_years",
    "parse_term_premium_mode_label",
    "term_benefit_timing_label_options",
    "term_contract_ui_config",
    "term_length_label_options",
    "term_metric_formatter",
    "term_premium_mode_label_options",
    "term_ui_config",
]
