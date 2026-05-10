"""SPIA adapter, metric formatter, and UI config (re-export shim).

Canonical import path for SPIA's UI-facing surface. The actual pricing
adapter, metric formatter, and ``ProductUIConfig`` instance live in
:mod:`product_registry`; this shim is the narrow access path that keeps
the consolidated :mod:`products.spia` package self-contained.

Notes
-----
The SPIA adapter, formatter, and UI config are read directly from
:mod:`product_registry` rather than re-implemented here. Any future
move of the adapter into this subpackage just changes the imports
below; no caller needs to update.
"""

from __future__ import annotations

from annuity_model.product_registry import (
    _SPIA_ADAPTER,
    ProductUIConfig,
    _spia_pricing_metrics,
)

# Resolve at import time so the module-level constants are stable refs
# (downstream `is`-comparisons in tests rely on this).
SPIA_ADAPTER = _SPIA_ADAPTER
spia_metric_formatter = _spia_pricing_metrics
spia_ui_config = ProductUIConfig(
    selected_info_message=None,
    projection_csv_filename="pricing_projection_spia.csv",
    recalc_workbook_filename="spia_recalc_model.xlsx",
)


__all__ = [
    "SPIA_ADAPTER",
    "spia_metric_formatter",
    "spia_ui_config",
]
