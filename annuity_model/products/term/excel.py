"""Term Life Excel build-spec + builder (re-export shim).

Implementation lives in :mod:`build_term_excel_workbook`.
"""

from __future__ import annotations

from build_term_excel_workbook import (
    TermExcelBuildSpec,
    build_term_workbook_from_spec,
    term_excel_spec_from_launcher,
)

__all__ = [
    "TermExcelBuildSpec",
    "build_term_workbook_from_spec",
    "term_excel_spec_from_launcher",
]
