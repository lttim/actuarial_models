"""RILA Excel build-spec + builder (re-export shim).

Implementation lives in :mod:`build_rila_excel_workbook`.
"""

from __future__ import annotations

from build_rila_excel_workbook import (
    RILAExcelBuildSpec,
    build_rila_workbook_from_spec,
    rila_excel_spec_from_launcher,
)

__all__ = [
    "RILAExcelBuildSpec",
    "build_rila_workbook_from_spec",
    "rila_excel_spec_from_launcher",
]
