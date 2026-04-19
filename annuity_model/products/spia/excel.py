"""SPIA Excel build-spec + builder (re-export shim).

Canonical import path for SPIA's Excel pipeline:

.. code-block:: python

    from products.spia.excel import (
        ExcelBuildSpec,
        excel_spec_from_launcher,
        build_workbook_from_spec,
    )

The actual implementation still lives in
:mod:`build_pricing_excel_workbook`. Snapshot dataclasses
(``ExcelPythonSnapshot``, ``MCExcelSnapshot``, ``ALMExcelSnapshot``) are
re-exported too because the spec consumes them.
"""

from __future__ import annotations

from build_pricing_excel_workbook import (
    ALMExcelSnapshot,
    ExcelBuildSpec,
    ExcelPythonSnapshot,
    MCExcelSnapshot,
    build_workbook_from_spec,
    excel_spec_from_launcher,
)

__all__ = [
    "ALMExcelSnapshot",
    "ExcelBuildSpec",
    "ExcelPythonSnapshot",
    "MCExcelSnapshot",
    "build_workbook_from_spec",
    "excel_spec_from_launcher",
]
