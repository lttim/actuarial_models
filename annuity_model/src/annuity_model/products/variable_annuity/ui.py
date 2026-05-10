"""VARIABLE_ANNUITY adapter, metric formatter, and UI config (re-export shim)."""

from __future__ import annotations

from annuity_model.product_registry import (
    _VA_ADAPTER,
    ProductUIConfig,
    _accumulation_pricing_metrics,
)

VARIABLE_ANNUITY_ADAPTER = _VA_ADAPTER
variable_annuity_metric_formatter = _accumulation_pricing_metrics
variable_annuity_ui_config = ProductUIConfig(
    selected_info_message="Variable Annuity (single premium): GMDB = max(AV, premium). Sub-account is deterministic CSV by default; Monte Carlo simulates GBM.",
    projection_csv_filename="pricing_projection_variable_annuity.csv",
    recalc_workbook_filename="variable_annuity_recalc_model.xlsx",
)


__all__ = [
    "VARIABLE_ANNUITY_ADAPTER",
    "variable_annuity_metric_formatter",
    "variable_annuity_ui_config",
]
