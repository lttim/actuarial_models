"""Smoke tests for :mod:`portfolio_runner`."""

from __future__ import annotations

from pathlib import Path

import numpy as np
import pytest

from annuity_model import pricing_projection as sp
from annuity_model import rila_projection as rp
from annuity_model import term_projection as tp
from annuity_model.inforce_io import load_policy_inputs_from_csv
from annuity_model.portfolio import PolicyInput, Portfolio, RunScenario
from annuity_model.portfolio_runner import run_portfolio
from annuity_model.pricing_scenario_materialize import (
    ANN_MODEL_ROOT,
    run_scenario_for_portfolio_policies,
)
from annuity_model.product_registry import ProductType


def _scenario() -> RunScenario:
    yc = sp.YieldCurve.from_flat_rate(0.04)
    ages = np.arange(0, 121, dtype=int)
    qx = np.full_like(ages, 0.02, dtype=float)
    mort = sp.MortalityTableQx(ages, qx)
    ex = sp.ExpenseAssumptions(0.0, 0.0, 0.0)
    return RunScenario(
        yield_curve=yc,
        mortality=mort,
        horizon_age=90,
        spread=0.0,
        valuation_year=None,
        expenses=ex,
        expenses_csv_path=sp.DEFAULT_EXPENSES_CSV,
        index_scenario_csv_path=None,
        expense_annual_inflation=0.0,
    )


def test_run_portfolio_parallel_matches_serial() -> None:
    scen = _scenario()
    pol = Portfolio(
        policies=(
            PolicyInput(ProductType.SPIA, sp.SPIAContract(65, "male", 50_000.0)),
            PolicyInput(
                ProductType.TERM_LIFE,
                tp.TermLifeContract(40, "female", 100_000.0, monthly_premium=50.0, term_years=15),
            ),
        )
    )
    a = run_portfolio(portfolio=pol, scenario=scen, max_workers=1)
    b = run_portfolio(portfolio=pol, scenario=scen, max_workers=2)
    assert len(a.policy_results) == len(b.policy_results)
    assert float(a.liability_path_total.expected_total_cashflows.sum()) == pytest.approx(
        float(b.liability_path_total.expected_total_cashflows.sum())
    )


def test_default_portfolio_scenario_rila_has_positive_outputs() -> None:
    """Regression: zero-expense portfolio scenarios priced RILA at SP=0 (degenerate)."""
    from annuity_model.portfolio_scenario import default_run_scenario

    scen = default_run_scenario(default_product_type=ProductType.RILA)
    pol = Portfolio(
        policies=(
            PolicyInput(
                ProductType.RILA,
                rp.RILAContract(
                    issue_age=55,
                    sex="female",
                    participation=0.8,
                    cap=0.07,
                    floor=0.0,
                    rider_fee_annual=0.01,
                ),
            ),
        )
    )
    res = run_portfolio(portfolio=pol, scenario=scen)
    rila_pr = next(pr for pr in res.policy_results if pr.product_type == ProductType.RILA)
    sp = float(getattr(rila_pr.pricing, "single_premium"))
    assert sp > 0.0
    cf_sum = float(np.sum(np.asarray(rila_pr.pricing.expected_total_cashflows, dtype=float)))
    assert cf_sum > 0.0


def test_run_portfolio_rila_degenerate_expense_context_error() -> None:
    scen = _scenario()
    pol = Portfolio(
        policies=(
            PolicyInput(
                ProductType.RILA,
                rp.RILAContract(
                    issue_age=55,
                    sex="female",
                    participation=0.8,
                    cap=0.07,
                    floor=0.0,
                    rider_fee_annual=0.01,
                ),
                policy_id="rila-bad",
            ),
        )
    )
    with pytest.raises(ValueError, match="policy_id='rila-bad'.*product_type='rila'"):
        run_portfolio(portfolio=pol, scenario=scen)


def test_run_portfolio_baseline_alm_matches_direct_liability_path() -> None:
    root = Path(__file__).resolve().parents[2]
    policies = load_policy_inputs_from_csv(root / "tests/data/inforce/example_v1/inforce.csv")
    sex_raw = str(getattr(policies[0].contract, "sex", "male")).strip().lower()
    sex = "female" if sex_raw == "female" else "male"
    scen = run_scenario_for_portfolio_policies({}, policies, sex=sex, repo_root=ANN_MODEL_ROOT)
    asm = sp.alm_engine_baseline_assumptions()
    res = run_portfolio(portfolio=Portfolio(policies=policies), scenario=scen, alm_assumptions=asm)
    assert res.alm_result is not None
    aum = sum(float(pr.pricing.single_premium) for pr in res.policy_results)  # type: ignore[attr-defined]
    direct = sp.run_alm_projection_from_liability_path(
        liability_path=res.liability_path_total,
        yield_curve=scen.yield_curve,
        spread=scen.spread,
        assumptions=asm,
        initial_asset_market_value=aum,
    )
    np.testing.assert_allclose(res.alm_result.surplus, direct.surplus, rtol=0, atol=1e-6)
    np.testing.assert_allclose(
        res.alm_result.funding_ratio, direct.funding_ratio, rtol=0, atol=1e-9
    )


def test_run_portfolio_two_spias_mixed_types() -> None:
    scen = _scenario()
    pol = Portfolio(
        policies=(
            PolicyInput(ProductType.SPIA, sp.SPIAContract(65, "male", 50_000.0)),
            PolicyInput(
                ProductType.TERM_LIFE,
                tp.TermLifeContract(40, "female", 100_000.0, monthly_premium=50.0, term_years=15),
            ),
        )
    )
    res = run_portfolio(portfolio=pol, scenario=scen)
    assert len(res.policy_results) == 2
    assert {pr.product_type for pr in res.policy_results} == {
        ProductType.SPIA,
        ProductType.TERM_LIFE,
    }
    assert len(res.rollups_by_product_type) == 2
    assert len(res.liability_path_total.expected_total_cashflows) >= 1
