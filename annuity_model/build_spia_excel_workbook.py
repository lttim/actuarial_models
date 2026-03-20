from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.chart import LineChart, Reference


BASE_DIR = Path(".")
OUT_XLSX = BASE_DIR / "spia_projection_model.xlsx"


def _load_data() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    yc = pd.read_csv(BASE_DIR / "treasury_zero_rate_curve_latest.csv")
    rp = pd.read_csv(BASE_DIR / "rp2014_male_healthy_annuitant_qx_2014.csv")
    mp = pd.read_csv(BASE_DIR / "mp2016_male_improvement_rates.csv")
    exp = pd.read_csv(BASE_DIR / "expenses_assumptions_us_placeholders.csv")
    return yc, rp, mp, exp


def _read_text(path: Path, fallback: str = "") -> str:
    try:
        return path.read_text(encoding="utf-8")
    except Exception:
        return fallback


def _write_inputs(ws, exp_df: pd.DataFrame) -> None:
    ws.title = "Inputs"
    ws["A1"] = "SPIA Inputs"
    ws["A1"].font = Font(bold=True, size=12)

    rows = [
        ("Issue Age", 65),
        ("Sex", "male"),
        ("Annual Benefit", 100000),
        ("Payment Frequency", 12),
        ("Valuation Year", 2025),
        ("Horizon Age", 110),
        ("Spread", 0.0),
    ]

    # Pull placeholders from CSV
    key_to_val = {k: v for k, v in zip(exp_df["key"], exp_df["value"])}
    rows.extend(
        [
            ("Policy Expense Dollars", float(key_to_val.get("policy_expense_dollars", 250))),
            ("Premium Expense Rate", float(key_to_val.get("premium_expense_rate", 0.01))),
            ("Monthly Expense Dollars", float(key_to_val.get("monthly_expense_dollars", 25))),
        ]
    )

    for i, (k, v) in enumerate(rows, start=3):
        ws[f"A{i}"] = k
        ws[f"B{i}"] = v

    ws["A15"] = "Derived"
    ws["A15"].font = Font(bold=True)
    ws["A16"] = "Monthly Benefit"
    ws["B16"] = "=B5/B6"
    ws["A17"] = "Projection Months"
    ws["B17"] = "=(B8-B3)*B6"

    ws["D3"] = "Notes"
    ws["D4"] = "All rates are decimals."
    ws["D5"] = "Valuation date interpreted as 12/31/Valuation Year."
    ws["D6"] = "Monthly expenses assumed paid while in-force."


def _write_simple_table(ws, title: str, df: pd.DataFrame) -> None:
    ws["A1"] = title
    ws["A1"].font = Font(bold=True, size=12)
    for c, col in enumerate(df.columns, start=1):
        ws.cell(row=3, column=c, value=col).font = Font(bold=True)
    for r, row in enumerate(df.itertuples(index=False), start=4):
        for c, val in enumerate(row, start=1):
            ws.cell(row=r, column=c, value=float(val) if isinstance(val, (int, float)) else val)


def _write_monthly_curve(ws, n_months: int, yield_rows: int) -> None:
    ws.title = "MonthlyCurve"
    ws["A1"] = "Monthly Interpolated Zero Curve"
    ws["A1"].font = Font(bold=True, size=12)

    headers = [
        "Month",
        "t_years",
        "BracketIndex",
        "LowerMat",
        "UpperMat",
        "LowerZero",
        "UpperZero",
        "InterpWeight",
        "ZeroRateInterpolated",
        "DiscountFactor",
    ]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=3, column=c, value=h).font = Font(bold=True)

    first = 4
    last = first + n_months - 1
    for r in range(first, last + 1):
        ws[f"A{r}"] = r - first + 1
        ws[f"B{r}"] = f"=A{r}/Inputs!$B$6"
        # Bracket index k such that m_k <= t <= m_{k+1}, with endpoint handling.
        ws[f"C{r}"] = (
            f"=IF(B{r}<=INDEX(YieldCurve!$A$4:$A${yield_rows},1),1,"
            f"IF(B{r}>=INDEX(YieldCurve!$A$4:$A${yield_rows},ROWS(YieldCurve!$A$4:$A${yield_rows})),"
            f"ROWS(YieldCurve!$A$4:$A${yield_rows})-1,"
            f"MATCH(B{r},YieldCurve!$A$4:$A${yield_rows},1)))"
        )
        ws[f"D{r}"] = f"=INDEX(YieldCurve!$A$4:$A${yield_rows},C{r})"
        ws[f"E{r}"] = f"=INDEX(YieldCurve!$A$4:$A${yield_rows},C{r}+1)"
        ws[f"F{r}"] = f"=INDEX(YieldCurve!$B$4:$B${yield_rows},C{r})"
        ws[f"G{r}"] = f"=INDEX(YieldCurve!$B$4:$B${yield_rows},C{r}+1)"
        ws[f"H{r}"] = f"=IF(E{r}=D{r},0,(B{r}-D{r})/(E{r}-D{r}))"
        ws[f"I{r}"] = (
            f"=IF(B{r}<=INDEX(YieldCurve!$A$4:$A${yield_rows},1),INDEX(YieldCurve!$B$4:$B${yield_rows},1),"
            f"IF(B{r}>=INDEX(YieldCurve!$A$4:$A${yield_rows},ROWS(YieldCurve!$A$4:$A${yield_rows})),"
            f"INDEX(YieldCurve!$B$4:$B${yield_rows},ROWS(YieldCurve!$B$4:$B${yield_rows})),"
            f"F{r}+H{r}*(G{r}-F{r})))"
        )
        ws[f"J{r}"] = f"=EXP(-(I{r}+Inputs!$B$9)*B{r})"


def _write_projection(ws, n_months: int) -> None:
    ws.title = "Projection"
    ws["A1"] = "SPIA Monthly Projection (Formula Driven)"
    ws["A1"].font = Font(bold=True, size=12)

    headers = [
        "Month",
        "t_years",
        "AttainedAge",
        "IntAge",
        "CalendarYear",
        "MonthlyBenefit",
        "MonthlyExpense",
        "BaseQx2014",
        "CumMP",
        "QxYear",
        "MuYear",
        "SurvInterval",
        "SurvivalToPay",
        "ZeroRateInterpolated",
        "DiscountFactor",
        "ExpBenefitCF",
        "ExpExpenseCF",
        "ExpTotalCF",
        "PVBenefitCF",
        "PVExpenseCF",
        "ReserveAfterPayment",
    ]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=3, column=c, value=h).font = Font(bold=True)

    first = 4
    last = first + n_months - 1

    for r in range(first, last + 1):
        ws[f"A{r}"] = r - first + 1
        ws[f"B{r}"] = f"=A{r}/Inputs!$B$6"
        ws[f"C{r}"] = f"=Inputs!$B$3+B{r}"
        ws[f"D{r}"] = f"=INT(C{r})"
        ws[f"E{r}"] = f"=Inputs!$B$7+1+INT((A{r}-1)/Inputs!$B$6)"
        ws[f"F{r}"] = "=Inputs!$B$16"
        ws[f"G{r}"] = "=Inputs!$B$12"
        ws[f"H{r}"] = f'=IFERROR(INDEX(RP2014_Qx!$B:$B,MATCH(D{r},RP2014_Qx!$A:$A,0)),"")'
        ws[f"I{r}"] = (
            f'=SUMIFS(MP2016_Long!$C:$C,MP2016_Long!$A:$A,D{r},'
            f'MP2016_Long!$B:$B,">=2014",MP2016_Long!$B:$B,"<"&E{r})'
        )
        ws[f"J{r}"] = f"=MIN(0.999999,MAX(0,H{r}*EXP(I{r})))"
        ws[f"K{r}"] = f"=-LN(1-J{r})"
        ws[f"L{r}"] = f"=EXP(-K{r}/Inputs!$B$6)"
        if r == first:
            ws[f"M{r}"] = f"=L{r}"
        else:
            ws[f"M{r}"] = f"=M{r-1}*L{r}"

        # Pull monthly-interpolated zero/discount from helper tab.
        ws[f"N{r}"] = f'=IFERROR(INDEX(MonthlyCurve!$I:$I,MATCH(A{r},MonthlyCurve!$A:$A,0)),"")'
        ws[f"O{r}"] = f'=IFERROR(INDEX(MonthlyCurve!$J:$J,MATCH(A{r},MonthlyCurve!$A:$A,0)),"")'
        ws[f"P{r}"] = f"=F{r}*M{r}"
        ws[f"Q{r}"] = f"=G{r}*M{r}"
        ws[f"R{r}"] = f"=P{r}+Q{r}"
        ws[f"S{r}"] = f"=P{r}*O{r}"
        ws[f"T{r}"] = f"=Q{r}*O{r}"

        if r == last:
            ws[f"U{r}"] = 0.0
        else:
            ws[f"U{r}"] = f"=SUMPRODUCT(R{r+1}:R{last},O{r+1}:O{last})/(M{r}*O{r})"

    # Summary block
    ws["W3"] = "Summary"
    ws["W3"].font = Font(bold=True)
    ws["W4"] = "PV Benefits"
    ws["X4"] = f"=SUM(S{first}:S{last})"
    ws["W5"] = "PV Monthly Expenses"
    ws["X5"] = f"=SUM(T{first}:T{last})"
    ws["W6"] = "Annuity Factor"
    ws["X6"] = f"=SUMPRODUCT(M{first}:M{last},O{first}:O{last})"
    ws["W7"] = "PV Monthly Total"
    ws["X7"] = f"=X4+X5"
    ws["W8"] = "Single Premium"
    ws["X8"] = "=(Inputs!$B$10+X7)/(1-Inputs!$B$11)"
    ws["W9"] = "Reserve at t=0"
    ws["X9"] = "=X7"

    # Reserve time-zero row
    ws["A2"] = "ReserveAtT0"
    ws["B2"] = 0
    ws["C2"] = "=Inputs!$B$3"
    ws["U2"] = "=X9"


def _write_dashboard(wb: Workbook, n_months: int) -> None:
    ws = wb.create_sheet("Dashboard")
    ws["A1"] = "SPIA Policy Projection Dashboard"
    ws["A1"].font = Font(bold=True, size=14)

    # Key policy and assumption inputs
    ws["A3"] = "Policy Inputs"
    ws["A3"].font = Font(bold=True)
    ws["A4"] = "Issue Age"
    ws["B4"] = "=Inputs!B3"
    ws["A5"] = "Sex"
    ws["B5"] = "=Inputs!B4"
    ws["A6"] = "Annual Benefit"
    ws["B6"] = "=Inputs!B5"
    ws["A7"] = "Valuation Year (12/31)"
    ws["B7"] = "=Inputs!B7"
    ws["A8"] = "Horizon Age"
    ws["B8"] = "=Inputs!B8"
    ws["A9"] = "Spread"
    ws["B9"] = "=Inputs!B9"

    ws["D3"] = "Expense Assumptions"
    ws["D3"].font = Font(bold=True)
    ws["D4"] = "Policy Expense ($)"
    ws["E4"] = "=Inputs!B10"
    ws["D5"] = "Premium Expense Rate"
    ws["E5"] = "=Inputs!B11"
    ws["D6"] = "Monthly Expense ($)"
    ws["E6"] = "=Inputs!B12"

    # Output summary
    ws["A11"] = "Pricing & Reserve Summary"
    ws["A11"].font = Font(bold=True)
    ws["A12"] = "Single Premium"
    ws["B12"] = "=Projection!X8"
    ws["A13"] = "PV Benefits"
    ws["B13"] = "=Projection!X4"
    ws["A14"] = "PV Monthly Expenses"
    ws["B14"] = "=Projection!X5"
    ws["A15"] = "PV Monthly Total"
    ws["B15"] = "=Projection!X7"
    ws["A16"] = "Annuity Factor"
    ws["B16"] = "=Projection!X6"
    ws["A17"] = "Economic Reserve at t=0"
    ws["B17"] = "=Projection!U2"

    # Formatting
    for cell in ("B6", "E4", "E6", "B12", "B13", "B14", "B15", "B17"):
        ws[cell].number_format = '#,##0.00'
    ws["E5"].number_format = "0.00%"
    ws["B16"].number_format = "0.000000"

    proj_start = 4
    proj_end = proj_start + n_months - 1

    # Chart 1: Survival by age
    chart_surv = LineChart()
    chart_surv.title = "Survival to Monthly Payment"
    chart_surv.y_axis.title = "P(alive)"
    chart_surv.x_axis.title = "Attained Age"
    data_surv = Reference(wb["Projection"], min_col=13, min_row=proj_start, max_row=proj_end)
    cats_age = Reference(wb["Projection"], min_col=3, min_row=proj_start, max_row=proj_end)
    chart_surv.add_data(data_surv, titles_from_data=False)
    chart_surv.set_categories(cats_age)
    chart_surv.height = 6
    chart_surv.width = 9
    ws.add_chart(chart_surv, "A20")

    # Chart 2: Expected monthly benefit and expense cashflows
    chart_cf = LineChart()
    chart_cf.title = "Expected Monthly Cashflows"
    chart_cf.y_axis.title = "Expected Cashflow ($)"
    chart_cf.x_axis.title = "Attained Age"
    data_cf = Reference(wb["Projection"], min_col=16, max_col=17, min_row=3, max_row=proj_end)
    chart_cf.add_data(data_cf, titles_from_data=True)
    chart_cf.set_categories(cats_age)
    chart_cf.height = 6
    chart_cf.width = 9
    ws.add_chart(chart_cf, "J20")

    # Chart 3: Economic reserve by age
    chart_res = LineChart()
    chart_res.title = "Economic Reserve"
    chart_res.y_axis.title = "Reserve ($)"
    chart_res.x_axis.title = "Attained Age"
    data_res = Reference(wb["Projection"], min_col=21, min_row=2, max_row=proj_end)
    cats_res_age = Reference(wb["Projection"], min_col=3, min_row=2, max_row=proj_end)
    chart_res.add_data(data_res, titles_from_data=False)
    chart_res.set_categories(cats_res_age)
    chart_res.height = 6
    chart_res.width = 18
    ws.add_chart(chart_res, "A36")

    # Section 2: checkpoint table for quick pricing memo snapshots
    ws["A54"] = "Checkpoint Metrics (Nearest Monthly Projection Point)"
    ws["A54"].font = Font(bold=True)
    cp_headers = ["Target Age", "Projection Row", "Survival", "Exp Benefit ($)", "Exp Expense ($)", "Reserve ($)"]
    for c, h in enumerate(cp_headers, start=1):
        ws.cell(row=55, column=c, value=h).font = Font(bold=True)

    checkpoint_ages = [65, 70, 75, 80, 85, 90, 100, 110]
    for i, age in enumerate(checkpoint_ages, start=56):
        ws[f"A{i}"] = age
        # Find checkpoint row in Projection:
        # - exact age if available
        # - otherwise first row if below first projection age
        # - otherwise previous row via approximate match
        ws[f"B{i}"] = (
            f"=IFERROR(MATCH(A{i},Projection!$C$4:$C${3+n_months},0)+3,"
            f"IF(A{i}<INDEX(Projection!$C$4:$C${3+n_months},1),4,"
            f"MATCH(A{i},Projection!$C$4:$C${3+n_months},1)+3))"
        )
        ws[f"C{i}"] = f"=INDEX(Projection!$M:$M,B{i})"
        ws[f"D{i}"] = f"=INDEX(Projection!$P:$P,B{i})"
        ws[f"E{i}"] = f"=INDEX(Projection!$Q:$Q,B{i})"
        ws[f"F{i}"] = f"=INDEX(Projection!$U:$U,B{i})"

    # Number formats for checkpoint section
    for r in range(56, 56 + len(checkpoint_ages)):
        ws[f"C{r}"].number_format = "0.000000"
        ws[f"D{r}"].number_format = "#,##0.00"
        ws[f"E{r}"].number_format = "#,##0.00"
        ws[f"F{r}"].number_format = "#,##0.00"


def _write_python_tab(wb: Workbook) -> None:
    ws = wb.create_sheet("Python_Runbook")
    ws["A1"] = "Python Runbook / Repro Steps"
    ws["A1"].font = Font(bold=True, size=14)

    lines: list[tuple[str, str]] = []
    lines.append(("Purpose", "This tab documents how to run and regenerate the SPIA Python model artifacts."))
    lines.append(("Working folder", "annuity_model"))
    lines.append(("Python version", "Use Python 3.11+ recommended."))
    lines.append(("Dependencies", "numpy, pandas, openpyxl, matplotlib"))
    lines.append(("", ""))
    lines.append(("Core commands", "From the annuity_model folder run:"))
    lines.append(("1", "python spia_projection.py"))
    lines.append(("2", "python illustrate_spia_projection.py"))
    lines.append(("3", "python build_spia_excel_workbook.py"))
    lines.append(("", ""))
    lines.append(("Inputs used by Python", "treasury_zero_rate_curve_latest.csv"))
    lines.append(("", "rp2014_mort_tab_rates_exposure.xlsx"))
    lines.append(("", "mp2016_rates.xlsx"))
    lines.append(("", "expenses_assumptions_us_placeholders.csv"))
    lines.append(("", ""))
    lines.append(("Outputs generated by Python", "spia_projection_model.xlsx"))
    lines.append(("", "illustrations/*.png"))
    lines.append(("", "rp2014_male_healthy_annuitant_qx_2014.csv"))
    lines.append(("", "mp2016_male_improvement_rates.csv"))
    lines.append(("", ""))
    lines.append(("Python in Excel note", "If your Excel has Python-in-Excel, you can use =PY() cells."))
    lines.append(("Example", '=PY("import pandas as pd; pd.read_csv(\'treasury_zero_rate_curve_latest.csv\').head()")'))
    lines.append(("Example", '=PY("from spia_projection import _example_usage; _example_usage()")'))
    lines.append(("", ""))
    lines.append(("Caveat", "Python-in-Excel runs in Microsoft sandbox; local file access may be restricted."))
    lines.append(("Recommendation", "Use terminal commands above for production reproducibility."))

    row = 3
    for k, v in lines:
        ws[f"A{row}"] = k
        ws[f"B{row}"] = v
        if k in {"Purpose", "Core commands", "Inputs used by Python", "Outputs generated by Python", "Python in Excel note"}:
            ws[f"A{row}"].font = Font(bold=True)
        row += 1

    ws["A35"] = "spia_projection.py (excerpt)"
    ws["A35"].font = Font(bold=True)
    spia_text = _read_text(BASE_DIR / "spia_projection.py", fallback="spia_projection.py not found")
    excerpt = "\n".join(spia_text.splitlines()[:140])
    ws["A36"] = excerpt

    ws["A95"] = "illustrate_spia_projection.py (excerpt)"
    ws["A95"].font = Font(bold=True)
    ill_text = _read_text(BASE_DIR / "illustrate_spia_projection.py", fallback="illustrate_spia_projection.py not found")
    excerpt2 = "\n".join(ill_text.splitlines()[:120])
    ws["A96"] = excerpt2

    # Widen key columns for readability
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 120


def build_workbook() -> Path:
    yc, rp, mp, exp = _load_data()

    wb = Workbook()
    ws_inputs = wb.active
    _write_inputs(ws_inputs, exp)

    ws_yc = wb.create_sheet("YieldCurve")
    _write_simple_table(ws_yc, "Treasury Zero Curve", yc)

    ws_rp = wb.create_sheet("RP2014_Qx")
    _write_simple_table(ws_rp, "RP-2014 Male Healthy Annuitant Base Qx (2014)", rp)

    ws_mp = wb.create_sheet("MP2016_Long")
    _write_simple_table(ws_mp, "MP-2016 Male Improvement Rates (Long)", mp)

    n_months = int((110 - 65) * 12)
    ws_monthly_curve = wb.create_sheet("MonthlyCurve")
    _write_monthly_curve(
        ws_monthly_curve,
        n_months=n_months,
        yield_rows=3 + len(yc),
    )
    ws_proj = wb.create_sheet("Projection")
    _write_projection(ws_proj, n_months=n_months)
    _write_dashboard(wb, n_months=n_months)
    _write_python_tab(wb)

    wb.save(OUT_XLSX)
    return OUT_XLSX


if __name__ == "__main__":
    out = build_workbook()
    print(f"Created workbook: {out.resolve()}")

