"""
excel_formula_sim_template.py
-----------------------------
Template: Python simulation of Excel formula logic.

Copy this file to your product's tests/parity/excel_formula_sim.py and implement
the methods to match your product's Excel formula logic.

Any change to Excel formula generators must be reflected here, and vice versa.
This file is the "executable specification" of the Excel formulas.
"""

from __future__ import annotations

import numpy as np


# ---------------------------------------------------------------------------
# Template: Selection / Disinvestment formula simulation
# ---------------------------------------------------------------------------


def excel_select_shortest_first(
    need: float,
    keys: np.ndarray,        # sort keys (e.g. remaining tenor)
    values: np.ndarray,      # amounts to select from (e.g. face values)
    discounts: np.ndarray,   # conversion factor (e.g. discount factor)
) -> tuple[float, np.ndarray]:
    """Simulate Excel's multi-pass shortest-first selection formula.

    Replicates the pattern:
        adjusted_key[k] = key[k] + (k+1)*EPSILON
        min_key = MIN(IF(value[k]>DEPLETED_THRESHOLD, adjusted_key[k], SENTINEL))
        select[k] = IF(AND(value[k]>DEPLETED_THRESHOLD,
                           ABS(adjusted_key[k]-min_key)<THRESHOLD),
                       MIN(value[k], need/discount[k]), 0)

    NOTE: THRESHOLD must be < EPSILON/2 to prevent double-selection.
    Adapt EPSILON and THRESHOLD to match your product's Excel formula.

    Parameters
    ----------
    need:       Amount to raise.
    keys:       Sort key per item (lower = selected first).
    values:     Amount available per item.
    discounts:  Conversion factor: cash raised = select[k] * discounts[k].

    Returns
    -------
    (remaining_need, values_after_selection)
    """
    # --- ADAPT THESE CONSTANTS TO YOUR PRODUCT ---
    EPSILON = 1e-9             # per-item epsilon increment
    THRESHOLD = 5e-10          # must be < EPSILON/2
    DEPLETED_THRESHOLD = 1e-9  # items below this are considered depleted
    SENTINEL = 999.0           # value for depleted items in min search
    N_PASSES = len(values) + 2  # maximum selection passes
    # --- END ADAPT ---

    values = np.asarray(values, dtype=float).copy()
    n = len(values)
    remaining = float(need)

    for _ in range(N_PASSES):
        if remaining <= 1e-9:
            break

        adjusted = np.array([
            keys[k] + (k + 1) * EPSILON if values[k] > DEPLETED_THRESHOLD else SENTINEL
            for k in range(n)
        ])
        min_key = float(np.min(adjusted))

        selections = np.zeros(n)
        for k in range(n):
            if values[k] > DEPLETED_THRESHOLD and abs(adjusted[k] - min_key) < THRESHOLD:
                selections[k] = min(values[k], remaining / max(discounts[k], 1e-15))

        raised = float(np.sum(selections * discounts))
        remaining = max(0.0, remaining - raised)
        values = np.maximum(0.0, values - selections)

    return remaining, values


# ---------------------------------------------------------------------------
# Template: Reinvestment / Allocation formula simulation
# ---------------------------------------------------------------------------


def excel_reinvest_pro_rata(
    surplus: float,
    values: np.ndarray,      # current holdings per bucket
    gaps: np.ndarray,        # amount below target per bucket (>0 = underfunded)
    discounts: np.ndarray,   # conversion factor (e.g. discount factor)
    tenors: np.ndarray,      # nominal tenor or reset value per bucket
) -> tuple[float, np.ndarray, np.ndarray]:
    """Simulate Excel's pro-rata reinvestment formula.

    Parameters
    ----------
    surplus:    Cash available to reinvest (must be > 0).
    values:     Current holdings (face, units, etc.) per bucket.
    gaps:       Positive gap per bucket (0 for overfunded buckets).
    discounts:  Conversion factor: face received = cash_invested / discount.
    tenors:     Reset value for each bucket after reinvestment.

    Returns
    -------
    (surplus_after, values_after, tenors_reset_after)
    """
    values = np.asarray(values, dtype=float).copy()
    tenors = np.asarray(tenors, dtype=float).copy()
    gap_sum = float(np.sum(np.maximum(gaps, 0.0)))

    if surplus <= 1e-6 or gap_sum <= 1e-9:
        return surplus, values, tenors

    for k in range(len(values)):
        if gaps[k] <= 0:
            continue
        split = gaps[k] / gap_sum
        invested = surplus * split
        if invested <= 0 or discounts[k] <= 1e-15:
            continue
        values[k] += invested / discounts[k]
        tenors[k] = tenors[k]  # reset to nominal (tenors array is already nominal)
        surplus -= invested

    return max(0.0, surplus), values, tenors
