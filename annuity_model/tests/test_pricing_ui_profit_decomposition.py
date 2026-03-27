from __future__ import annotations

from types import SimpleNamespace

import numpy as np

import pricing_projection as sp
import term_projection as tp
from pricing_ui import _build_profit_decomposition_rows, _build_profit_waterfall_chart_df
from product_registry import ProductType


def test_term_profit_decomposition_excludes_spia_indexation_language() -> None:
    result = SimpleNamespace(
        expected_benefit_cashflows=np.array([0.0, 120.0, 80.0], dtype=float),
        pv_benefit=180.0,
        pv_monthly_expenses=-70.0,
        single_premium=110.0,
    )
    contract = tp.TermLifeContract(issue_age=40, sex="male", death_benefit=100_000.0, monthly_premium=50.0)

    rows, _ = _build_profit_decomposition_rows(
        res=result,
        contract=contract,
        expenses=None,
        product_type=ProductType.TERM_LIFE,
    )
    labels = [label for label, _, _ in rows]
    assert all("indexation" not in label.lower() for label in labels)


def test_term_profit_decomposition_reconciles_to_net_pv() -> None:
    result = SimpleNamespace(
        expected_benefit_cashflows=np.array([100.0, 75.0, 25.0], dtype=float),
        pv_benefit=170.0,
        pv_monthly_expenses=-60.0,
        single_premium=110.0,
    )
    contract = tp.TermLifeContract(issue_age=40, sex="female", death_benefit=80_000.0, monthly_premium=40.0)

    rows, _ = _build_profit_decomposition_rows(
        res=result,
        contract=contract,
        expenses=None,
        product_type=ProductType.TERM_LIFE,
    )
    baseline = float(rows[0][1])
    deltas = float(rows[1][1]) + float(rows[2][1])
    final_total = float(rows[-1][1])
    assert np.isclose(baseline + deltas, final_total)


def test_spia_profit_decomposition_uses_product_design_effect_label() -> None:
    result = SimpleNamespace(
        months=np.arange(1, 4, dtype=int),
        survival_to_payment=np.array([0.99, 0.98, 0.97], dtype=float),
        discount_factors=np.array([0.995, 0.99, 0.985], dtype=float),
        pv_benefit=240.0,
        pv_monthly_expenses=12.0,
        single_premium=260.0,
    )
    contract = sp.SPIAContract(issue_age=65, sex="male", benefit_annual=1200.0)
    expenses = sp.ExpenseAssumptions(policy_expense_dollars=5.0, premium_expense_rate=0.03, monthly_expense_dollars=1.0)

    rows, _ = _build_profit_decomposition_rows(
        res=result,
        contract=contract,
        expenses=expenses,
        product_type=ProductType.SPIA,
    )
    labels = [label for label, _, _ in rows]
    assert "Benefit design effect (e.g., indexation)" in labels


def test_waterfall_walk_term_like_decreases_and_reconciliation() -> None:
    rows = [
        ("Undiscounted expected claims", 200.0, True),
        ("Discounting effect", -30.0, False),
        ("Policyholder premium PV (funding)", -60.0, False),
        ("Net PV (claims - premiums)", 110.0, True),
    ]
    df = _build_profit_waterfall_chart_df(rows)
    assert df.iloc[0]["lo"] == 0.0 and df.iloc[0]["hi"] == 200.0 and df.iloc[0]["bar_color"] == "Total"
    assert df.iloc[1]["delta"] == -30.0
    assert df.iloc[1]["start"] == 200.0 and df.iloc[1]["end"] == 170.0 and df.iloc[1]["bar_color"] == "Decrease"
    assert df.iloc[2]["start"] == 170.0 and df.iloc[2]["end"] == 110.0 and df.iloc[2]["bar_color"] == "Decrease"
    assert df.iloc[3]["lo"] == 0.0 and df.iloc[3]["hi"] == 110.0 and df.iloc[3]["bar_color"] == "Total"


def test_waterfall_walk_increase_step_and_signed_delta() -> None:
    rows = [
        ("Start", 100.0, True),
        ("Gain", 25.0, False),
        ("End", 125.0, True),
    ]
    df = _build_profit_waterfall_chart_df(rows)
    assert df.iloc[1]["delta"] == 25.0
    assert df.iloc[1]["start"] == 100.0 and df.iloc[1]["end"] == 125.0
    assert df.iloc[1]["bar_color"] == "Increase"


def test_spia_profit_decomposition_rows_reconcile_to_single_premium() -> None:
    result = SimpleNamespace(
        months=np.arange(1, 4, dtype=int),
        survival_to_payment=np.array([0.99, 0.98, 0.97], dtype=float),
        discount_factors=np.array([0.995, 0.99, 0.985], dtype=float),
        pv_benefit=240.0,
        pv_monthly_expenses=12.0,
        single_premium=260.0,
    )
    contract = sp.SPIAContract(issue_age=65, sex="male", benefit_annual=1200.0)
    expenses = sp.ExpenseAssumptions(policy_expense_dollars=5.0, premium_expense_rate=0.03, monthly_expense_dollars=1.0)

    rows, _ = _build_profit_decomposition_rows(
        res=result,
        contract=contract,
        expenses=expenses,
        product_type=ProductType.SPIA,
    )
    anchor = float(rows[0][1])
    bridge = sum(float(v) for _, v, is_total in rows if not is_total)
    final = float(rows[-1][1])
    assert np.isclose(anchor + bridge, final)
