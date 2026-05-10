"""MYGA Excel build-spec + builder (re-export shim).

Implementation lives in :mod:`build_myga_excel_workbook`.
"""

from __future__ import annotations

from annuity_model.build_myga_excel_workbook import (
    MYGAExcelBuildSpec,
    build_myga_workbook_from_spec,
    myga_excel_spec_from_launcher,
)

__all__ = [
    "MYGAExcelBuildSpec",
    "build_myga_workbook_from_spec",
    "myga_excel_spec_from_launcher",
]
