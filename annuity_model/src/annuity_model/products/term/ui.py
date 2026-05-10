"""Term Life adapter, metric formatter, and UI config (re-export shim).

Re-exports the Term-Life-specific UI surface (selectbox label parsers,
``TermContractUIConfig`` instance, etc.) under the canonical
``products.term.ui`` namespace.
"""

from __future__ import annotations

from annuity_model.product_registry import (
    _TERM_ADAPTER,
    ProductUIConfig,
    _term_pricing_metrics,
    get_term_contract_ui_config,
    parse_term_benefit_timing_label,
    parse_term_length_label_to_years,
    parse_term_premium_mode_label,
    term_benefit_timing_label_options,
    term_length_label_options,
    term_premium_mode_label_options,
)

TERM_ADAPTER = _TERM_ADAPTER
term_metric_formatter = _term_pricing_metrics
term_ui_config = ProductUIConfig(
    selected_info_message="Term Life (20Y) is enabled with deterministic pricing. Monte Carlo is not available in this release.",
    projection_csv_filename="pricing_projection_term_life.csv",
    recalc_workbook_filename="term_life_recalc_model.xlsx",
)
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
