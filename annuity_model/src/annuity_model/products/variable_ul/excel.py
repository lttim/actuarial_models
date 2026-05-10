"""VARIABLE_UL Excel build-spec + builder (re-export shim).

Implementation lives in :mod:`build_vul_excel_workbook`.
"""

from __future__ import annotations

from annuity_model.build_vul_excel_workbook import (
    VULExcelBuildSpec,
    build_vul_workbook_from_spec,
    vul_excel_spec_from_launcher,
)

__all__ = [
    "VULExcelBuildSpec",
    "build_vul_workbook_from_spec",
    "vul_excel_spec_from_launcher",
]
