"""UNIVERSAL_LIFE adapter, metric formatter, and UI config (re-export shim)."""

from __future__ import annotations

from annuity_model.product_registry import (
    _UL_ADAPTER,
    ProductUIConfig,
    _life_single_premium_metrics,
)

UNIVERSAL_LIFE_ADAPTER = _UL_ADAPTER
universal_life_metric_formatter = _life_single_premium_metrics
universal_life_ui_config = ProductUIConfig(
    selected_info_message="Universal Life (single premium): monthly cycle of credit -> COI -> expense charge. Type A death benefit.",
    projection_csv_filename="pricing_projection_universal_life.csv",
    recalc_workbook_filename="universal_life_recalc_model.xlsx",
)


__all__ = [
    "UNIVERSAL_LIFE_ADAPTER",
    "universal_life_metric_formatter",
    "universal_life_ui_config",
]
