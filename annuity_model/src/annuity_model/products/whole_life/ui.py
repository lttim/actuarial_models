"""WHOLE_LIFE adapter, metric formatter, and UI config (re-export shim)."""

from __future__ import annotations

from annuity_model.product_registry import (
    _WL_ADAPTER,
    ProductUIConfig,
    _wl_pricing_metrics,
)

WHOLE_LIFE_ADAPTER = _WL_ADAPTER
whole_life_metric_formatter = _wl_pricing_metrics
whole_life_ui_config = ProductUIConfig(
    selected_info_message="Whole Life (single premium): premium solved as PV of benefits, mortality from CSO 2017 Ultimate placeholder.",
    projection_csv_filename="pricing_projection_whole_life.csv",
    recalc_workbook_filename="whole_life_recalc_model.xlsx",
)


__all__ = [
    "WHOLE_LIFE_ADAPTER",
    "whole_life_metric_formatter",
    "whole_life_ui_config",
]
