"""UNIVERSAL_LIFE Excel build-spec + builder (re-export shim).

Implementation lives in :mod:`build_ul_excel_workbook`.
"""

from __future__ import annotations

from build_ul_excel_workbook import (
    ULExcelBuildSpec,
    build_ul_workbook_from_spec,
    ul_excel_spec_from_launcher,
)

__all__ = [
    "ULExcelBuildSpec",
    "build_ul_workbook_from_spec",
    "ul_excel_spec_from_launcher",
]
