"""WHOLE_LIFE Excel build-spec + builder (re-export shim).

Implementation lives in :mod:`build_wl_excel_workbook`.
"""

from __future__ import annotations

from annuity_model.build_wl_excel_workbook import (
    WLExcelBuildSpec,
    build_wl_workbook_from_spec,
    wl_excel_spec_from_launcher,
)

__all__ = [
    "WLExcelBuildSpec",
    "build_wl_workbook_from_spec",
    "wl_excel_spec_from_launcher",
]
