"""MYGA adapter, metric formatter, and UI config (re-export shim)."""

from __future__ import annotations

from annuity_model.product_registry import (
    _MYGA_ADAPTER,
    ProductUIConfig,
    _accumulation_pricing_metrics,
)

MYGA_ADAPTER = _MYGA_ADAPTER
myga_metric_formatter = _accumulation_pricing_metrics
myga_ui_config = ProductUIConfig(
    selected_info_message="MYGA (multi-year guaranteed annuity): single premium accumulates at the declared rate for the guarantee period.",
    projection_csv_filename="pricing_projection_myga.csv",
    recalc_workbook_filename="myga_recalc_model.xlsx",
)


__all__ = [
    "MYGA_ADAPTER",
    "myga_metric_formatter",
    "myga_ui_config",
]
