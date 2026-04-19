"""Smoke tests for :mod:`portfolio_runner`."""

from __future__ import annotations

import numpy as np
import pytest

import pricing_projection as sp
import term_projection as tp
from portfolio import PolicyInput, Portfolio, RunScenario
from portfolio_runner import run_portfolio
from product_registry import ProductType


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
    assert {pr.product_type for pr in res.policy_results} == {ProductType.SPIA, ProductType.TERM_LIFE}
    assert len(res.rollups_by_product_type) == 2
    assert len(res.liability_path_total.expected_total_cashflows) >= 1
