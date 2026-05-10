"""VARIABLE_ANNUITY Excel build-spec + builder (re-export shim).

Implementation lives in :mod:`build_va_excel_workbook`.
"""

from __future__ import annotations

from annuity_model.build_va_excel_workbook import (
    VAExcelBuildSpec,
    build_va_workbook_from_spec,
    va_excel_spec_from_launcher,
)

__all__ = [
    "VAExcelBuildSpec",
    "build_va_workbook_from_spec",
    "va_excel_spec_from_launcher",
]
