"""Canonical inforce with one default row per :class:`ProductType`."""

from __future__ import annotations

from pathlib import Path

import pytest

from annuity_model import pricing_projection as sp
from annuity_model.inforce_io import load_policy_inputs_from_csv
from annuity_model.portfolio import Portfolio
from annuity_model.portfolio_runner import run_portfolio
from annuity_model.pricing_scenario_materialize import (
    ANN_MODEL_ROOT,
    run_scenario_for_portfolio_policies,
)
from annuity_model.product_registry import ProductType


def test_all_products_default_inforce_loads_ten_distinct_types() -> None:
    root = Path(__file__).resolve().parents[2]
    policies = load_policy_inputs_from_csv(
        root / "tests/data/inforce/all_products_default_v1/inforce.csv"
    )
    types = {p.product_type for p in policies}
    assert len(policies) == 10
    assert types == set(ProductType)


def test_all_products_default_portfolio_runs_with_baseline_alm() -> None:
    root = Path(__file__).resolve().parents[2]
    policies = load_policy_inputs_from_csv(
        root / "tests/data/inforce/all_products_default_v1/inforce.csv"
    )
    scen = run_scenario_for_portfolio_policies({}, policies, sex="male", repo_root=ANN_MODEL_ROOT)
    asm = sp.alm_engine_baseline_assumptions()
    res = run_portfolio(portfolio=Portfolio(policies=policies), scenario=scen, alm_assumptions=asm)
    assert res.alm_result is not None
    assert len(res.rollups_by_product_type) == 10
    aum = sum(float(pr.pricing.single_premium) for pr in res.policy_results)  # type: ignore[attr-defined]
    assert aum > 0.0
    direct = sp.run_alm_projection_from_liability_path(
        liability_path=res.liability_path_total,
        yield_curve=scen.yield_curve,
        spread=scen.spread,
        assumptions=asm,
        initial_asset_market_value=aum,
    )
    assert float(res.alm_result.duration_gap) == pytest.approx(
        float(direct.duration_gap), rel=0, abs=1e-9
    )
