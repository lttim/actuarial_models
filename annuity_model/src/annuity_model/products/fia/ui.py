"""FIA adapter, metric formatter, and UI config (re-export shim)."""

from __future__ import annotations

from annuity_model.product_registry import (
    _FIA_ADAPTER,
    ProductUIConfig,
    _accumulation_pricing_metrics,
)

FIA_ADAPTER = _FIA_ADAPTER
fia_metric_formatter = _accumulation_pricing_metrics
fia_ui_config = ProductUIConfig(
    selected_info_message="FIA (fixed indexed annuity): annual point-to-point credit with cap, floor, and participation. Floor 0 by default.",
    projection_csv_filename="pricing_projection_fia.csv",
    recalc_workbook_filename="fia_recalc_model.xlsx",
)


__all__ = [
    "FIA_ADAPTER",
    "fia_metric_formatter",
    "fia_ui_config",
]
