"""Golden snapshot for canonical 5-policy mixed portfolio (example inforce)."""

from __future__ import annotations

import json
import os
from pathlib import Path

import pytest

from inforce_io import load_policy_inputs_from_csv
from portfolio import Portfolio
from portfolio_runner import run_portfolio
from portfolio_summary import portfolio_result_to_summary_dict
from pricing_scenario_materialize import ANN_MODEL_ROOT, run_scenario_for_portfolio_policies

pytestmark = pytest.mark.parity

GOLDEN_PATH = (
    Path(__file__).resolve().parents[1] / "golden" / "portfolio" / "portfolio_5policy.json"
)


def _current_snapshot() -> dict[str, object]:
    root = Path(__file__).resolve().parents[3]
    policies = load_policy_inputs_from_csv(root / "tests/data/inforce/example_v1/inforce.csv")
    sex = "female" if str(policies[0].contract.sex).lower() == "female" else "male"
    scen = run_scenario_for_portfolio_policies({}, policies, sex=sex, repo_root=ANN_MODEL_ROOT)
    res = run_portfolio(portfolio=Portfolio(policies=policies), scenario=scen)
    return {
        "summary": portfolio_result_to_summary_dict(res),
        "n_months_total_path": len(res.liability_path_total.expected_total_cashflows),
        "by_type_months": {
            pt.value: len(p.expected_total_cashflows)
            for pt, p in res.rollups_by_product_type.items()
        },
    }


def test_portfolio_5policy_golden_matches() -> None:
    if os.environ.get("UPDATE_GOLDEN_PORTFOLIO") == "1":
        GOLDEN_PATH.parent.mkdir(parents=True, exist_ok=True)
        GOLDEN_PATH.write_text(json.dumps(_current_snapshot(), indent=2) + "\n", encoding="utf-8")
        pytest.skip("Wrote golden; re-run without UPDATE_GOLDEN_PORTFOLIO=1.")

    expected = json.loads(GOLDEN_PATH.read_text(encoding="utf-8"))
    got = _current_snapshot()
    assert got == expected
