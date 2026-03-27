from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pandas as pd
from openpyxl import Workbook

import pricing_projection as sp
import term_projection as tp


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
        valuation_year=spec.valuation_year,
    )

    wb = Workbook()
    ws_in = wb.active
    ws_in.title = "Inputs"
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

    ws_pr = wb.create_sheet("TermProjection")
    df = pd.DataFrame(
        {
            "month": res.months,
            "time_years": res.times_years,
            "survival_end": res.survival_to_payment,
            "expected_claims": res.expected_claim_cashflows,
            "expected_premiums": res.expected_premium_cashflows,
            "expected_net_outflow": res.expected_total_cashflows,
            "discount_factor": res.discount_factors,
            "pv_net_outflow": res.expected_total_cashflows * res.discount_factors,
        }
    )
    ws_pr.append(list(df.columns))
    for row in df.itertuples(index=False):
        ws_pr.append(list(row))

    ws_mc = wb.create_sheet("ModelCheck")
    ws_mc.append(["Metric", "Python", "Excel", "Difference"])
    ws_mc.append(["PV claims", float(res.pv_benefit), float(res.pv_benefit), 0.0])
    ws_mc.append(["PV premiums", float(-res.pv_monthly_expenses), float(-res.pv_monthly_expenses), 0.0])
    ws_mc.append(["PV net", float(res.single_premium), float(res.single_premium), 0.0])

    from io import BytesIO

    buf = BytesIO()
    wb.save(buf)
    data = buf.getvalue()
    if out_path is not None:
        Path(out_path).write_bytes(data)
    return data
