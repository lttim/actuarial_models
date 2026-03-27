from __future__ import annotations

from pathlib import Path
from typing import Any

from build_pricing_excel_workbook import (
    ALMExcelSnapshot,
    ExcelBuildSpec,
    ExcelPythonSnapshot,
    MCExcelSnapshot,
    build_workbook_from_spec,
)
from build_term_excel_workbook import TermExcelBuildSpec, build_term_workbook_from_spec
from product_registry import ProductType
import pricing_projection as sp


def build_product_workbook(
    *,
    product_type: ProductType,
    spec: Any,
    out_path: str | Path | None = None,
    python_snapshot: ExcelPythonSnapshot | None = None,
    mc_snapshot: MCExcelSnapshot | None = None,
    alm_snapshot: ALMExcelSnapshot | None = None,
    alm_assumptions: sp.ALMAssumptions | None = None,
) -> bytes:
    if product_type == ProductType.SPIA:
        if not isinstance(spec, ExcelBuildSpec):
            raise TypeError("SPIA workbook builder requires ExcelBuildSpec.")
        return build_workbook_from_spec(
            spec,
            out_path=out_path,
            python_snapshot=python_snapshot,
            mc_snapshot=mc_snapshot,
            alm_snapshot=alm_snapshot,
            alm_assumptions=alm_assumptions,
        )
    if product_type == ProductType.TERM_LIFE:
        if not isinstance(spec, TermExcelBuildSpec):
            raise TypeError("Term workbook builder requires TermExcelBuildSpec.")
        return build_term_workbook_from_spec(
            spec,
            out_path=out_path,
            alm_snapshot=alm_snapshot,
            alm_assumptions=alm_assumptions,
        )
    raise NotImplementedError(f"Workbook builder is not implemented for product '{product_type.value}'.")
