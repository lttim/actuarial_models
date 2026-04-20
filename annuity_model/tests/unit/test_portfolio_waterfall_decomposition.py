"""Portfolio profit waterfall row logic (homogeneous vs mixed)."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from inforce_io import load_policy_inputs_from_csv, load_policy_inputs_from_csv_from_dataframe
from portfolio import Portfolio
from portfolio_runner import run_portfolio
from pricing_run_form_state import default_inforce_scratch_row
from pricing_scenario_materialize import ANN_MODEL_ROOT, run_scenario_for_portfolio_policies
from pricing_ui import (
    _build_portfolio_profit_decomposition_rows_for_policy_results,
    _build_portfolio_profit_decomposition_rows,
    _build_profit_decomposition_rows,
    _merge_profit_waterfall_row_sets,
)
from product_registry import ProductType


def test_homogeneous_spia_waterfall_rows_sum_per_policy_components() -> None:
    row_a = default_inforce_scratch_row(ProductType.SPIA)
    row_a["policy_id"] = "wa"
    row_b = {**default_inforce_scratch_row(ProductType.SPIA), "policy_id": "wb", "issue_age": 70}
    policies = load_policy_inputs_from_csv_from_dataframe(pd.DataFrame([row_a, row_b]))
    scen = run_scenario_for_portfolio_policies({}, policies, sex="male", repo_root=ANN_MODEL_ROOT)
    res = run_portfolio(portfolio=Portfolio(policies=policies), scenario=scen)
    rows_port, _cap = _build_portfolio_profit_decomposition_rows(res, scen.expenses)
    per_pol = [
        _build_profit_decomposition_rows(
            res=pr.pricing,  # type: ignore[arg-type]
            contract=pr.contract,
            expenses=scen.expenses,
            product_type=ProductType.SPIA,
        )[0]
        for pr in res.policy_results
    ]
    merged = _merge_profit_waterfall_row_sets(per_pol)
    assert len(merged) == len(rows_port)
    for a, b in zip(merged, rows_port, strict=True):
        assert a[0] == b[0] and a[2] == b[2]
        assert a[1] == pytest.approx(b[1], rel=0, abs=1e-6)


def test_mixed_book_generic_bridge_matches_scalar_sums() -> None:
    root = Path(__file__).resolve().parents[2]
    policies = load_policy_inputs_from_csv(root / "tests/data/inforce/example_v1/inforce.csv")
    sex_raw = str(getattr(policies[0].contract, "sex", "male")).strip().lower()
    sex = "female" if sex_raw == "female" else "male"
    scen = run_scenario_for_portfolio_policies({}, policies, sex=sex, repo_root=ANN_MODEL_ROOT)
    res = run_portfolio(portfolio=Portfolio(policies=policies), scenario=scen)
    rows, _cap = _build_portfolio_profit_decomposition_rows(res, scen.expenses)
    sum_pv_b = sum(float(getattr(pr.pricing, "pv_benefit")) for pr in res.policy_results)
    sum_pv_m = sum(float(getattr(pr.pricing, "pv_monthly_expenses")) for pr in res.policy_results)
    sum_sp = sum(float(getattr(pr.pricing, "single_premium")) for pr in res.policy_results)
    n = len(res.policy_results)
    issue = (
        float(scen.expenses.policy_expense_dollars) * n
        if scen.expenses is not None
        else 0.0
    )
    assert rows[0] == ("PV benefits (portfolio sum)", sum_pv_b, True)
    assert rows[1][0] == "PV monthly cashflow component (portfolio sum)"
    assert rows[1][1] == pytest.approx(sum_pv_m)
    assert rows[2] == ("Issue expense (portfolio sum)", issue, False)
    assert rows[3][0] == "Modeled net premium / value (portfolio sum)" and rows[3][2] is True
    assert rows[3][1] == pytest.approx(sum_sp, rel=0, abs=1e-6)


def test_single_type_selection_from_mixed_book_matches_type_subset() -> None:
    root = Path(__file__).resolve().parents[2]
    policies = load_policy_inputs_from_csv(root / "tests/data/inforce/example_v1/inforce.csv")
    sex_raw = str(getattr(policies[0].contract, "sex", "male")).strip().lower()
    sex = "female" if sex_raw == "female" else "male"
    scen = run_scenario_for_portfolio_policies({}, policies, sex=sex, repo_root=ANN_MODEL_ROOT)
    res = run_portfolio(portfolio=Portfolio(policies=policies), scenario=scen)
    spia_only = tuple(pr for pr in res.policy_results if pr.product_type == ProductType.SPIA)
    rows_subset, _ = _build_portfolio_profit_decomposition_rows_for_policy_results(
        spia_only, scen.expenses
    )
    rows_expected, _ = _build_profit_decomposition_rows(
        res=spia_only[0].pricing,  # type: ignore[arg-type]
        contract=spia_only[0].contract,
        expenses=scen.expenses,
        product_type=ProductType.SPIA,
    )
    assert len(rows_subset) == len(rows_expected)
    for a, b in zip(rows_subset, rows_expected, strict=True):
        assert a[0] == b[0] and a[2] == b[2]
        assert a[1] == pytest.approx(b[1], rel=0, abs=1e-6)
