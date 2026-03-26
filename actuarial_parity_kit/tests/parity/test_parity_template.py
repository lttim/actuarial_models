"""
test_parity_template.py
-----------------------
Template: Cross-engine parity tests for an actuarial product model.

Copy this to your product's tests/parity/test_parity.py and adapt:
1. Replace imports with your engine's module names.
2. Implement _run_python_sim() and _run_excel_sim() for your product.
3. Add golden scenario parameters in GOLDEN_SCENARIOS.
4. Add regression tests for any historical bugs.

Run with:
    pytest tests/parity/ -v

All tests must pass at 0.00 discrepancy before any merge to main.
"""

from __future__ import annotations

import sys
from pathlib import Path

import numpy as np
import pytest

ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(ROOT))

# Replace with your product's imports:
# import my_product_engine as eng
# from tests.parity.excel_formula_sim import excel_select_shortest_first

# ---------------------------------------------------------------------------
# Tolerances (fill from docs/model_parity_contract.md)
# ---------------------------------------------------------------------------

TOL_PRIMARY = 1e-4   # primary output (e.g. account value, reserve)
TOL_FACTOR = 1e-10   # dimensionless factors (e.g. discount factor)

# ---------------------------------------------------------------------------
# Golden scenarios
# ---------------------------------------------------------------------------

GOLDEN_SCENARIOS = [
    # id, description, param1, param2, expected_discrepancy
    # ("baseline", "Flat 3% rate, no selection needed", 0.03, 2_000_000, 0.0),
    # ("light_selection", "Flat 4% rate, mild cash deficit", 0.04, 1_800_000, 0.0),
    # ("tie_break", "Nominal tie at step N", 0.037, 2_222_430, 0.0),
]


# ---------------------------------------------------------------------------
# Simulation helpers — ADAPT THESE TO YOUR PRODUCT
# ---------------------------------------------------------------------------


def _run_python_sim(params) -> list:
    """Run N-step simulation using the Python engine. Return state per step."""
    raise NotImplementedError("Implement for your product")


def _run_excel_sim(params) -> list:
    """Run N-step simulation using the Excel formula simulation. Return state per step."""
    raise NotImplementedError("Implement for your product")


# ---------------------------------------------------------------------------
# Tie-break unit tests — ADAPT AND KEEP
# ---------------------------------------------------------------------------


def test_selection_tie_break_lowest_index_wins():
    """When two items have equal sort keys, the lower-indexed is always selected first.

    ADAPT: Replace with your product's selection function and relevant state.
    """
    # Example structure — replace with your actual selection function:
    # keys = np.array([0.25, 0.25, 1.0, 3.0])  # items 0 and 1 tied
    # values = np.array([100_000.0, 200_000.0, 300_000.0, 400_000.0])
    # discounts = np.ones(4)
    # remaining, result = excel_select_shortest_first(5_000.0, keys, values, discounts)
    # assert result[0] < values[0]          # item 0 must be selected
    # np.testing.assert_allclose(result[1], values[1], atol=TOL_PRIMARY)  # item 1 untouched
    pytest.skip("Implement for your product (see template comments)")


@pytest.mark.regression
def test_no_double_selection_near_equal_keys():
    """Regression template: when two items have keys differing by ~1e-16, only one is selected.

    ADAPT: Build the floating-point accumulation scenario specific to your product.
    This is the most important regression test — it prevents the double-sell class of bug.
    """
    # Construct accumulated vs reset sort keys:
    # key_accumulated = start_value
    # for _ in range(N_steps): key_accumulated -= delta
    # key_reset = nominal_value
    # Assert abs(key_accumulated - key_reset) < 1e-14  # fp noise only
    # Assert selection picks only the lower-indexed item
    pytest.skip("Implement the floating-point accumulation scenario for your product")


# ---------------------------------------------------------------------------
# Full N-step parity tests — ADAPT
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("scenario_id,description,params,expected", GOLDEN_SCENARIOS)
def test_full_step_parity(scenario_id, description, params, expected):
    """Full step-by-step parity: Python engine vs Excel formula simulation.

    Asserts agreement at every step, not just the final output.
    """
    py_results = _run_python_sim(params)
    xl_results = _run_excel_sim(params)

    assert len(py_results) == len(xl_results), "Result lengths must match"
    for step, (py, xl) in enumerate(zip(py_results, xl_results)):
        np.testing.assert_allclose(
            py, xl, atol=TOL_PRIMARY,
            err_msg=f"[{scenario_id}] Mismatch at step {step+1}: Python={py:.6f}, Excel={xl:.6f}",
        )


@pytest.mark.regression
def test_golden_scenario_zero_final_discrepancy():
    """Regression: final output discrepancy must be exactly zero for all golden scenarios.

    ADAPT: Run your specific bug-scenario and assert zero discrepancy.
    Preserved permanently so any reversion of the fix is immediately caught.
    """
    pytest.skip("Implement: run the specific scenario that caused the historical discrepancy")
