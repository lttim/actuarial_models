"""Advisory benchmark: portfolio runner wall-clock on canonical inforce."""

from __future__ import annotations

import time
from pathlib import Path

import pytest

from annuity_model.inforce_io import load_policy_inputs_from_csv
from annuity_model.portfolio import Portfolio
from annuity_model.portfolio_runner import run_portfolio
from annuity_model.pricing_scenario_materialize import run_scenario_for_portfolio_policies


@pytest.mark.slow
def test_portfolio_five_policy_wall_clock() -> None:
    root = Path(__file__).resolve().parents[2]
    policies = load_policy_inputs_from_csv(root / "tests/data/inforce/example_v1/inforce.csv")
    sex = "female" if str(policies[0].contract.sex).lower() == "female" else "male"
    scen = run_scenario_for_portfolio_policies({}, policies, sex=sex)
    t0 = time.perf_counter()
    run_portfolio(portfolio=Portfolio(policies=policies), scenario=scen)
    elapsed = time.perf_counter() - t0
    assert elapsed < 30.0, f"portfolio run took {elapsed:.2f}s (expected <30s)"
