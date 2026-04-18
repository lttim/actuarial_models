"""Public surface for Excel-builder helpers shared across products.

Historically the Term and RILA builders reached into the SPIA builder for
``_write_alm_projection_sheet``, ``_write_model_check_sheet``,
``alm_excel_downsample_snapshot`` and friends. The leading underscore was a
lie -- these helpers are part of the cross-product Excel-build contract --
and the cross-imports made the dependency direction unobvious.

This shim re-exports the same callables under explicit public names. New
builders SHOULD import from here. The original ``_`` names in
:mod:`build_pricing_excel_workbook` are kept temporarily for
backwards-compatibility but should be considered private; remove them once
all callers move over.

Phase 2 deliverable; Phase 3 will fold these into a real
``BaseProductBuilder`` class with an overridable ``liability_layout``.
"""

from __future__ import annotations

# ALM ladder + downsampler (the SPIA builder owns the canonical impl)
from alm_excel_ladder import ALM_ENGINE_SHEET, write_alm_engine_sheet
from build_pricing_excel_workbook import (
    ALM_ENGINE_STEP_MONTHS,
    ALM_EXCEL_PATH_MONTH_CAP,
    ALM_PROJECTION_FIRST_DATA_ROW,
    LIABILITY_SHEET_NAME,
    ALMExcelSnapshot,
    ExcelPythonSnapshot,
    alm_excel_downsample_snapshot,
    alm_excel_truncate_snapshot,
    inject_alm_projection_formula_cached_values,
)

# Public aliases for the historically-`_private` helpers. Builders MUST use
# these names so a future internal rename does not silently break them.
from build_pricing_excel_workbook import _write_alm_projection_sheet as write_alm_projection_sheet
from build_pricing_excel_workbook import _write_model_check_sheet as write_model_check_sheet

# Layouts + validator: convenience re-exports so a builder needs only one
# import line to pull in the entire shared toolkit.
from excel_workbook_validator import validate_workbook, validate_workbook_or_raise
from liability_layouts import LIABILITY_LAYOUTS, LiabilityLayout, liability_layout_for

__all__ = [
    "ALM_ENGINE_SHEET",
    "ALM_ENGINE_STEP_MONTHS",
    "ALM_EXCEL_PATH_MONTH_CAP",
    "ALM_PROJECTION_FIRST_DATA_ROW",
    "ALMExcelSnapshot",
    "ExcelPythonSnapshot",
    "LIABILITY_LAYOUTS",
    "LIABILITY_SHEET_NAME",
    "LiabilityLayout",
    "alm_excel_downsample_snapshot",
    "alm_excel_truncate_snapshot",
    "inject_alm_projection_formula_cached_values",
    "liability_layout_for",
    "validate_workbook",
    "validate_workbook_or_raise",
    "write_alm_engine_sheet",
    "write_alm_projection_sheet",
    "write_model_check_sheet",
]
