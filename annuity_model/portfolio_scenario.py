"""Default :class:`portfolio.RunScenario` for portfolio CLI / UI smoke tests."""

from __future__ import annotations

import numpy as np

import pricing_projection as sp
from portfolio import RunScenario


def default_run_scenario(
    *,
    horizon_age: int = 95,
    flat_rate: float = 0.04,
    qx_flat: float = 0.02,
) -> RunScenario:
    """Flat yield, flat q_x mortality, zero expenses — matches many unit tests."""
    yc = sp.YieldCurve.from_flat_rate(flat_rate)
    ages = np.arange(0, 121, dtype=int)
    qx = np.full_like(ages, qx_flat, dtype=float)
    mort = sp.MortalityTableQx(ages, qx)
    ex = sp.ExpenseAssumptions(0.0, 0.0, 0.0)
    return RunScenario(
        yield_curve=yc,
        mortality=mort,
        horizon_age=horizon_age,
        spread=0.0,
        valuation_year=None,
        expenses=ex,
        expenses_csv_path=sp.DEFAULT_EXPENSES_CSV,
        index_scenario_csv_path=None,
        expense_annual_inflation=0.0,
    )


__all__ = ["default_run_scenario"]
