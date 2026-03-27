from __future__ import annotations

import math
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Literal

import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.worksheet.worksheet import Worksheet

import pricing_projection as sp
import term_projection as tp

TERM_PROJ_MAX_ROWS = 600

SHEET_INPUTS = "Inputs"
SHEET_ZERO = "ZeroCurve"
SHEET_QX = "QxTable"
SHEET_MTH_QX = "MortalMonthly"
SHEET_PROJ = "TermProjection"
SHEET_MC = "ModelCheck"

_IN_ROW_ISSUE_AGE = 4
_IN_ROW_DEATH_BEN = 6
_IN_ROW_PREMIUM = 7
_IN_ROW_TERM_Y = 8
_IN_ROW_HORIZON = 11
_IN_ROW_SPREAD = 13
_IN_ROW_NMONTHS = 17


def _in_addr(col: str, row: int) -> str:
    return f"{SHEET_INPUTS}!${col}${row}"


def _n_months_cell() -> str:
    return _in_addr("B", _IN_ROW_NMONTHS)


def _write_zero_curve_sheet(wb: Workbook, yc: sp.YieldCurve, spread_cell: str) -> int:
    ws = wb.create_sheet(SHEET_ZERO)
    ws["A1"] = "maturity_years"
    ws["B1"] = "zero_rate"
    ws["C1"] = "ln_discount_node"
    mats = np.asarray(yc.maturities_years, dtype=float)
    zeros = np.asarray(yc.zero_rates, dtype=float)
    n = int(mats.shape[0])
    for i in range(n):
        r = 2 + i
        ws.cell(row=r, column=1, value=float(mats[i]))
        ws.cell(row=r, column=2, value=float(zeros[i]))
        ws.cell(row=r, column=3, value=f"=-(B{r}+{spread_cell})*A{r}")
    return n


def _write_qx_table_sheet(wb: Workbook, mort: sp.MortalityTableQx) -> int:
    ws = wb.create_sheet(SHEET_QX)
    ws["A1"] = "age"
    ws["B1"] = "qx"
    ages = np.asarray(mort.ages, dtype=int)
    qx = np.asarray(mort.qx, dtype=float)
    n = int(ages.shape[0])
    for i in range(n):
        r = 2 + i
        ws.cell(row=r, column=1, value=int(ages[i]))
        ws.cell(row=r, column=2, value=float(qx[i]))
    return n


def _fill_mortal_monthly_rpmp(
    ws_m: Worksheet,
    *,
    mort: sp.MortalityTableRP2014MP2016,
    issue_age: int,
    valuation_year: int,
    n_months: int,
) -> None:
    dt = 1.0 / 12.0
    for k in range(1, n_months + 1):
        r = 1 + k
        m_index = k - 1
        age_start = issue_age + m_index * dt
        age_int = int(math.floor(age_start))
        calendar_year_start = valuation_year + 1 + (m_index // 12)
        qxv = float(
            mort.qx_at_int_age_and_calendar_year(age_int=age_int, calendar_year=calendar_year_start)
        )
        ws_m.cell(row=r, column=1, value=f"=IF(ROW()-1>{_n_months_cell()},\"\",ROW()-1)")
        ws_m.cell(row=r, column=2, value=int(age_int))
        ws_m.cell(row=r, column=3, value=int(calendar_year_start))
        ws_m.cell(row=r, column=4, value=float(qxv))
    last_data = 1 + n_months
    last_cap = 1 + TERM_PROJ_MAX_ROWS
    for r in range(last_data + 1, last_cap + 1):
        ws_m.cell(row=r, column=1, value=f"=IF(ROW()-1>{_n_months_cell()},\"\",ROW()-1)")
        for c in (2, 3, 4):
            ws_m.cell(row=r, column=c, value="")


@dataclass(frozen=True)
class TermExcelBuildSpec:
    contract: tp.TermLifeContract
    yield_curve: sp.YieldCurve
    mortality: sp.MortalityTableQx | sp.MortalityTableRP2014MP2016
    horizon_age: int
    spread: float
    valuation_year: int
    expenses: sp.ExpenseAssumptions
    yield_mode_label: str
    mortality_mode_label: str
    expense_mode_label: str


def term_excel_spec_from_launcher(
    *,
    contract: tp.TermLifeContract,
    yield_curve: sp.YieldCurve,
    mortality: sp.MortalityTableQx | sp.MortalityTableRP2014MP2016,
    horizon_age: int,
    spread: float,
    valuation_year: int,
    expenses: sp.ExpenseAssumptions,
    yield_mode_label: str,
    mortality_mode_label: str,
    expense_mode_label: str,
) -> TermExcelBuildSpec:
    return TermExcelBuildSpec(
        contract=contract,
        yield_curve=yield_curve,
        mortality=mortality,
        horizon_age=int(horizon_age),
        spread=float(spread),
        valuation_year=int(valuation_year),
        expenses=expenses,
        yield_mode_label=str(yield_mode_label),
        mortality_mode_label=str(mortality_mode_label),
        expense_mode_label=str(expense_mode_label),
    )


def _qx_lookup_expr(acell: str, mode: Literal["qx_table", "mortal_monthly"]) -> str:
    clamp_inner = "MIN(MAX({inner},0),0.999999)"
    if mode == "qx_table":
        inner = (
            f"INDEX({SHEET_QX}!$B$2:$B$50000,"
            f"MATCH(INT({_in_addr('B', _IN_ROW_ISSUE_AGE)}+({acell}-1)/12),"
            f"{SHEET_QX}!$A$2:$A$50000,0))"
        )
        return clamp_inner.format(inner=inner)
    inner = (
        f"INDEX({SHEET_MTH_QX}!$D$2:$D$50000,"
        f"MATCH({acell},{SHEET_MTH_QX}!$A$2:$A$50000,0))"
    )
    return clamp_inner.format(inner=inner)


def _survival_end_formula(r: int, acell: str, mode: Literal["qx_table", "mortal_monthly"]) -> str:
    qx_e = _qx_lookup_expr(acell, mode)
    p_m = f"EXP(-(-LN(1-{qx_e}))/12)"
    if r == 2:
        return f"=IF({acell}=\"\",\"\",{p_m})"
    prev = f"C{r-1}"
    return f"=IF({acell}=\"\",\"\",{prev}*{p_m})"


def _discount_factor_formula(r: int, n_pts: int) -> str:
    spr = _in_addr("B", _IN_ROW_SPREAD)
    acell = f"A{r}"
    tcell = f"B{r}"
    zr = f"{SHEET_ZERO}!$A$2:$A${1 + n_pts}"
    zv = f"{SHEET_ZERO}!$B$2:$B${1 + n_pts}"
    lndf = f"{SHEET_ZERO}!$C$2:$C${1 + n_pts}"
    last_a = f"INDEX({zr},{n_pts})"
    last_z = f"INDEX({zv},{n_pts})"
    return (
        f"=IF({acell}=\"\",\"\","
        f"IF({tcell}<={SHEET_ZERO}!$A$2,EXP(-({SHEET_ZERO}!$B$2+{spr})*{tcell}),"
        f"IF({tcell}>={last_a},EXP(-({last_z}+{spr})*{tcell}),"
        f"EXP(INDEX({lndf},MATCH({tcell},{zr},1))"
        f"+(INDEX({lndf},MATCH({tcell},{zr},1)+1)-INDEX({lndf},MATCH({tcell},{zr},1)))"
        f"*({tcell}-INDEX({zr},MATCH({tcell},{zr},1)))"
        f"/(INDEX({zr},MATCH({tcell},{zr},1)+1)-INDEX({zr},MATCH({tcell},{zr},1)))"
        f"))))"
    )


def build_term_workbook_from_spec(
    spec: TermExcelBuildSpec,
    *,
    out_path: str | Path | None = None,
) -> bytes:
    res = tp.price_term_life_level_monthly(
        contract=spec.contract,
        yield_curve=spec.yield_curve,
        mortality=spec.mortality,
        horizon_age=spec.horizon_age,
        spread=spec.spread,
        valuation_year=spec.valuation_year
        if not isinstance(spec.mortality, sp.MortalityTableQx)
        else None,
    )

    wb = Workbook()
    ws_in = wb.active
    ws_in.title = SHEET_INPUTS
    ws_in["A1"] = "Term Life Inputs"
    rows = [
        ("Product", "20Y Level Term Life"),
        ("Issue age", spec.contract.issue_age),
        ("Sex", spec.contract.sex),
        ("Death benefit", spec.contract.death_benefit),
        ("Monthly premium", spec.contract.monthly_premium),
        ("Term years", spec.contract.term_years),
        ("Benefit timing", spec.contract.benefit_timing),
        ("Premium mode", spec.contract.premium_mode),
        ("Horizon age", spec.horizon_age),
        ("Valuation year", spec.valuation_year),
        ("Spread", spec.spread),
        ("Yield mode", spec.yield_mode_label),
        ("Mortality mode", spec.mortality_mode_label),
        ("Expense mode", spec.expense_mode_label),
    ]
    for i, (k, v) in enumerate(rows, start=3):
        ws_in[f"A{i}"] = k
        ws_in[f"B{i}"] = v
    nm = (
        f"=MIN(MAX(1,ROUND(({_in_addr('B', _IN_ROW_HORIZON)}"
        f"-{_in_addr('B', _IN_ROW_ISSUE_AGE)})*12,0)),"
        f"{_in_addr('B', _IN_ROW_TERM_Y)}*12)"
    )
    ws_in[f"A{_IN_ROW_NMONTHS}"] = "Model months (formula)"
    ws_in[f"B{_IN_ROW_NMONTHS}"] = nm
    ws_in[f"B{_IN_ROW_NMONTHS}"].number_format = "0"

    spread_cell = _in_addr("B", _IN_ROW_SPREAD)
    n_curve = _write_zero_curve_sheet(wb, spec.yield_curve, spread_cell)

    if isinstance(spec.mortality, sp.MortalityTableQx):
        _write_qx_table_sheet(wb, spec.mortality)
        mort_mode: Literal["qx_table", "mortal_monthly"] = "qx_table"
    else:
        mort_mode = "mortal_monthly"
        ws_m = wb.create_sheet(SHEET_MTH_QX)
        ws_m["A1"] = "month"
        ws_m["B1"] = "age_int"
        ws_m["C1"] = "calendar_year_start"
        ws_m["D1"] = "qx_annual"
        _fill_mortal_monthly_rpmp(
            ws_m,
            mort=spec.mortality,
            issue_age=spec.contract.issue_age,
            valuation_year=int(spec.valuation_year),
            n_months=int(res.months.shape[0]),
        )

    ws_pr = wb.create_sheet(SHEET_PROJ)
    hdr = (
        "month",
        "time_years",
        "survival_end",
        "survival_start",
        "month_death_prob",
        "expected_claims",
        "expected_premiums",
        "expected_net_outflow",
        "discount_factor",
        "pv_net_outflow",
    )
    for c, h in enumerate(hdr, start=1):
        ws_pr.cell(row=1, column=c, value=h)
        ws_pr.cell(row=1, column=c).font = Font(bold=True)

    nm_ref = _n_months_cell()
    last_cap_row = 1 + TERM_PROJ_MAX_ROWS
    ben = _in_addr("B", _IN_ROW_DEATH_BEN)
    prem = _in_addr("B", _IN_ROW_PREMIUM)
    for r in range(2, last_cap_row + 1):
        a = f"A{r}"
        ws_pr.cell(row=r, column=1, value=f"=IF(ROW()-1>{nm_ref},\"\",ROW()-1)")
        ws_pr.cell(row=r, column=2, value=f"=IF({a}=\"\",\"\",{a}/12)")
        ws_pr.cell(row=r, column=3, value=_survival_end_formula(r, a, mort_mode))
        if r == 2:
            d_surv_start = f"=IF({a}=\"\",\"\",1)"
        else:
            d_surv_start = f"=IF({a}=\"\",\"\",C{r-1})"
        ws_pr.cell(row=r, column=4, value=d_surv_start)
        ws_pr.cell(
            row=r,
            column=5,
            value=f"=IF({a}=\"\",\"\",MAX(0,MIN(1,D{r}-C{r})))",
        )
        ws_pr.cell(
            row=r,
            column=6,
            value=(
                f"=IF({a}=\"\",0,IF(MOD({a},12)=0,{ben}*SUM(OFFSET(E{r},-11,0,12,1)),0))"
            ),
        )
        ws_pr.cell(row=r, column=7, value=f"=IF({a}=\"\",0,{prem}*D{r})")
        ws_pr.cell(row=r, column=8, value=f"=IF({a}=\"\",0,F{r}-G{r})")
        ws_pr.cell(row=r, column=9, value=_discount_factor_formula(r, n_curve))
        ws_pr.cell(row=r, column=10, value=f"=IF({a}=\"\",0,H{r}*I{r})")

    for c in range(1, 11):
        ws_pr.cell(row=1, column=c).number_format = "General"
    money_cols = (6, 7, 8, 10)
    for r in range(2, last_cap_row + 1):
        for c in money_cols:
            ws_pr.cell(row=r, column=c).number_format = "#,##0.00"
        for c in (2, 3, 4, 5, 9):
            ws_pr.cell(row=r, column=c).number_format = "0.000000"

    lr = last_cap_row
    tp_ref = SHEET_PROJ
    sum_claims = f"=SUMPRODUCT({tp_ref}!$F$2:$F${lr},{tp_ref}!$I$2:$I${lr})"
    sum_prem = f"=SUMPRODUCT({tp_ref}!$G$2:$G${lr},{tp_ref}!$I$2:$I${lr})"
    sum_net = f"=SUM({tp_ref}!$J$2:$J${lr})"

    ws_mc = wb.create_sheet(SHEET_MC)
    ws_mc["A1"] = "Python snapshot vs Excel (Inputs → curves → TermProjection)"
    ws_mc["A1"].font = Font(bold=True, size=12)
    ws_mc["A2"] = (
        "Column B is the Python snapshot at export. Column C aggregates the TermProjection "
        "grid (all formulas); column D should be ~0 after a full recalc. Edit Inputs, "
        "ZeroCurve, QxTable (or MortalMonthly q_x for RP/MP runs) and recalculate."
    )
    ws_mc.merge_cells("A2:D2")

    hdr_row = 4
    headers = ("Metric", "Python snapshot", "Excel (formula)", "Difference (Excel − Python)")
    for c, h in enumerate(headers, start=1):
        ws_mc.cell(row=hdr_row, column=c, value=h)
        ws_mc.cell(row=hdr_row, column=c).font = Font(bold=True)

    pricing_rows: list[tuple[str, float, str]] = [
        ("PV claims", float(res.pv_benefit), sum_claims),
        ("PV premiums", float(-res.pv_monthly_expenses), sum_prem),
        ("PV net (claims − premiums)", float(res.single_premium), sum_net),
    ]
    row_idx = hdr_row + 1
    for label, val, xls_formula in pricing_rows:
        ws_mc.cell(row=row_idx, column=1, value=label)
        ws_mc.cell(row=row_idx, column=2, value=float(val))
        ws_mc.cell(row=row_idx, column=3, value=xls_formula)
        ws_mc.cell(row=row_idx, column=4, value=f"=C{row_idx}-B{row_idx}")
        for col in (2, 3, 4):
            ws_mc.cell(row=row_idx, column=col).number_format = "#,##0.00"
        row_idx += 1

    buf = BytesIO()
    wb.save(buf)
    data = buf.getvalue()
    if out_path is not None:
        Path(out_path).write_bytes(data)
    return data
