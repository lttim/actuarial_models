"""VARIABLE_UL adapter, metric formatter, and UI config (re-export shim)."""

from __future__ import annotations

from annuity_model.product_registry import (
    _VUL_ADAPTER,
    ProductUIConfig,
    _life_single_premium_metrics,
)

VARIABLE_UL_ADAPTER = _VUL_ADAPTER
variable_ul_metric_formatter = _life_single_premium_metrics
variable_ul_ui_config = ProductUIConfig(
    selected_info_message="Variable UL (VUL): UL mechanics with sub-account return as credit (deterministic CSV or GBM Monte Carlo).",
    projection_csv_filename="pricing_projection_variable_ul.csv",
    recalc_workbook_filename="variable_ul_recalc_model.xlsx",
)


__all__ = [
    "VARIABLE_UL_ADAPTER",
    "variable_ul_metric_formatter",
    "variable_ul_ui_config",
]
