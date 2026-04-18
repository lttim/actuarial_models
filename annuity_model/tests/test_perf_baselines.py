"""Performance baseline gates (P4).

Pinned baselines are intentionally generous: this file exists to catch
*regressions* (a 25% slowdown), not to micro-optimise. Bump the limits in a
PR that explains the underlying methodology change.

Skipped if `pytest-benchmark` is not installed.
"""

from __future__ import annotations

import io

import numpy as np
import pytest

pytest.importorskip("pytest_benchmark")

import pricing_projection as sp
import rila_projection as rp
from build_pricing_excel_workbook import build_workbook_from_spec, excel_spec_from_launcher
from build_rila_excel_workbook import build_rila_workbook_from_spec, rila_excel_spec_from_launcher
from excel_workbook_validator import validate_workbook_or_raise
from openpyxl import load_workbook

pytestmark = [pytest.mark.slow]

# Baselines are upper bounds in seconds, on macos-arm64 / Python 3.12.
# CI currently asserts the test ran without exceeding these (pytest-benchmark
# auto-fails on regressions if `--benchmark-compare-fail=mean:25%` is supplied
# in CI invocation; otherwise the values surface as informational).
SPIA_BUILD_BUDGET_S = 1.5
RILA_BUILD_BUDGET_S = 1.0
VALIDATOR_BUDGET_S = 1.5


def _spia_inputs():
    contract = sp.SPIAContract(issue_age=65, sex="male", benefit_annual=120_000.0)
    yc = sp.YieldCurve.from_flat_rate(0.045)
    ages = np.arange(0, 121, dtype=int)
    qx = np.clip(0.008 + ages * 2e-5, 1e-6, 0.4)
    mort = sp.MortalityTableQx(ages, qx)
    expenses = sp.ExpenseAssumptions(0.0, 0.0, 25.0)
    return contract, yc, mort, expenses


def _rila_inputs():
    contract = rp.RILAContract(
        issue_age=55, sex="male", participation=0.85, cap=0.09, floor=-0.02, rider_fee_annual=0.008
    )
    yc = sp.YieldCurve.from_flat_rate(0.035)
    ages = np.arange(0, 121, dtype=int)
    qx = np.clip(0.008 + ages * 2e-5, 1e-6, 0.4)
    mort = sp.MortalityTableQx(ages, qx)
    expenses = sp.ExpenseAssumptions(0.0, 0.0, 25.0)
    rng = np.random.default_rng(42)
    n_months = int(round((90 - contract.issue_age) * 12))
    levels = 100.0 * np.cumprod(1.0 + rng.normal(0.004, 0.02, size=n_months))
    return contract, yc, mort, expenses, levels


def _spia_spec(horizon_age: int = 90):
    contract, yc, mort, expenses = _spia_inputs()
    res = sp.price_spia_single_premium(
        contract=contract, yield_curve=yc, mortality=mort,
        horizon_age=horizon_age, spread=0.0,
        expenses=expenses, expense_annual_inflation=0.0,
    )
    spec = excel_spec_from_launcher(
        contract=contract, yield_curve=yc, mortality=mort, horizon_age=horizon_age,
        spread=0.0, valuation_year=2025, expenses=expenses,
        yield_mode_label="flat", mortality_mode_label="synthetic", expense_mode_label="manual",
        index_s0=float(res.index_s0),
        index_levels_at_payment=res.index_level_at_payment,
        expense_annual_inflation=0.0,
    )
    return spec


def test_spia_workbook_build_under_budget(benchmark) -> None:
    spec = _spia_spec()
    blob = benchmark(build_workbook_from_spec, spec)
    assert isinstance(blob, (bytes, bytearray))
    assert benchmark.stats.stats.mean < SPIA_BUILD_BUDGET_S, (
        f"SPIA build mean {benchmark.stats.stats.mean:.3f}s "
        f"exceeded budget {SPIA_BUILD_BUDGET_S}s"
    )


def test_rila_workbook_build_under_budget(benchmark) -> None:
    contract, yc, mort, expenses, levels = _rila_inputs()
    spec = rila_excel_spec_from_launcher(
        contract=contract, yield_curve=yc, mortality=mort, horizon_age=90,
        spread=0.0, valuation_year=2025, expenses=expenses,
        yield_mode_label="flat", mortality_mode_label="synthetic",
        expense_mode_label="manual",
        index_s0=100.0, index_levels_at_payment=levels, expense_annual_inflation=0.0,
    )

    blob = benchmark(build_rila_workbook_from_spec, spec)
    assert isinstance(blob, (bytes, bytearray))
    assert benchmark.stats.stats.mean < RILA_BUILD_BUDGET_S, (
        f"RILA build mean {benchmark.stats.stats.mean:.3f}s "
        f"exceeded budget {RILA_BUILD_BUDGET_S}s"
    )


def test_validator_under_budget(benchmark) -> None:
    spec = _spia_spec()
    blob = build_workbook_from_spec(spec)
    wb = load_workbook(io.BytesIO(blob))

    benchmark(validate_workbook_or_raise, wb)
    assert benchmark.stats.stats.mean < VALIDATOR_BUDGET_S, (
        f"Validator mean {benchmark.stats.stats.mean:.3f}s "
        f"exceeded budget {VALIDATOR_BUDGET_S}s"
    )
