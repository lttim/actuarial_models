"""Build a portfolio workbook (per-policy + by-type + total liability rollups + ModelCheck)."""

from __future__ import annotations

import io
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from excel_workbook_validator import validate_workbook_or_raise
from portfolio import PortfolioResult, ProductTypeRollupScalars
from product_registry import ProductType


def _sorted_product_types(rollups: dict[ProductType, object]) -> tuple[ProductType, ...]:
    return tuple(sorted(rollups, key=lambda p: p.value))


def build_portfolio_workbook_bytes(res: PortfolioResult) -> bytes:
    """Create an .xlsx with liability rollups; ``ModelCheck`` asserts sum(types)==total per month."""
    wb = Workbook()
    # --- Inputs ---
    ws_in = wb.active
    ws_in.title = "Inputs"
    ws_in["A1"] = "Portfolio workbook (v1)"
    ws_in["A1"].font = Font(bold=True, size=12)
    ws_in["A2"] = "Liability cashflows are Python literals; ModelCheck uses Excel formulas."

    # --- PolicyRegister ---
    ws_pr = wb.create_sheet("PolicyRegister")
    headers = ("policy_id", "product_type", "single_premium")
    for c, h in enumerate(headers, start=1):
        ws_pr.cell(row=1, column=c, value=h).font = Font(bold=True)
    for r, pr in enumerate(res.policy_results, start=2):
        ws_pr.cell(row=r, column=1, value=pr.policy_id)
        ws_pr.cell(row=r, column=2, value=pr.product_type.value)
        prem = getattr(pr.pricing, "single_premium", None)
        ws_pr.cell(row=r, column=3, value=float(prem) if prem is not None else None)

    # --- ProductTypeRollups ---
    ws_rt = wb.create_sheet("ProductTypeRollups")
    ws_rt["A1"] = "product_type"
    ws_rt["B1"] = "policy_count"
    ws_rt["C1"] = "sum_single_premium"
    ws_rt["D1"] = "sum_undiscounted_cf"
    for c in range(1, 5):
        ws_rt.cell(row=1, column=c).font = Font(bold=True)
    r = 2
    for pt in _sorted_product_types(dict(res.rollups_by_product_type)):
        scal: ProductTypeRollupScalars = res.product_type_scalar_rollups[pt]
        path = res.rollups_by_product_type[pt]
        sum_cf = float(path.expected_total_cashflows.sum()) if len(path.expected_total_cashflows) else 0.0
        ws_rt.cell(row=r, column=1, value=pt.value)
        ws_rt.cell(row=r, column=2, value=scal.policy_count)
        ws_rt.cell(
            row=r,
            column=3,
            value=scal.sum_single_premium if scal.sum_single_premium is not None else None,
        )
        ws_rt.cell(
            row=r,
            column=4,
            value=scal.sum_undiscounted_cashflows if scal.sum_undiscounted_cashflows is not None else sum_cf,
        )
        r += 1

    # --- LiabilityAggregate ---
    ws_la = wb.create_sheet("LiabilityAggregate")
    total = res.liability_path_total
    n = len(total.expected_total_cashflows)
    types = _sorted_product_types(dict(res.rollups_by_product_type))
    ws_la["A1"] = "month"
    ws_la["B1"] = "t_years"
    ws_la["C1"] = "total_cf"
    for j, pt in enumerate(types, start=4):
        ws_la.cell(row=1, column=j, value=f"cf_{pt.value}").font = Font(bold=True)
    for c in range(1, 4):
        ws_la.cell(row=1, column=c).font = Font(bold=True)

    last_type_col = 3 + len(types)
    last_letter = get_column_letter(last_type_col)

    for i in range(n):
        rr = 2 + i
        ws_la.cell(row=rr, column=1, value=i + 1)
        ws_la.cell(row=rr, column=2, value=float(total.times_years[i]))
        ws_la.cell(row=rr, column=3, value=float(total.expected_total_cashflows[i]))
        for j, pt in enumerate(types, start=4):
            cf_j = res.rollups_by_product_type[pt].expected_total_cashflows
            v = float(cf_j[i]) if i < len(cf_j) else 0.0
            ws_la.cell(row=rr, column=j, value=v)

    # --- ModelCheck ---
    ws_mc = wb.create_sheet("ModelCheck")
    ws_mc["A1"] = "month"
    ws_mc["B1"] = "rollup_minus_total"
    ws_mc["A1"].font = Font(bold=True)
    ws_mc["B1"].font = Font(bold=True)
    for i in range(n):
        rr = 2 + i
        ws_mc.cell(row=rr, column=1, value=i + 1)
        first_t = get_column_letter(4)
        ws_mc.cell(
            row=rr,
            column=2,
            value=f"=SUM(LiabilityAggregate!{first_t}{rr}:{last_letter}{rr})-LiabilityAggregate!C{rr}",
        )

    # --- README ---
    ws_rm = wb.create_sheet("README")
    ws_rm["A1"] = "Portfolio v1 workbook"
    ws_rm["A2"] = "Per-policy pricing is seriatim in Python; this file rolls up liability CF only."
    ws_rm["A3"] = "ALM is not embedded here (v1); run ALM from the app or CLI on the aggregate path."

    validate_workbook_or_raise(wb)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_portfolio_workbook_to_path(res: PortfolioResult, out_path: str | Path) -> Path:
    p = Path(out_path)
    p.write_bytes(build_portfolio_workbook_bytes(res))
    return p


__all__ = ["build_portfolio_workbook_bytes", "build_portfolio_workbook_to_path"]
