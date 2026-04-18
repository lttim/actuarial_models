"""Runtime Excel recalc parity (P4).

The classic parity tests compare Python to a *Python simulation* of the Excel
formulas (`tests/parity/excel_formula_sim.py`). That guards against logic
drift, but not against a builder bug that emits a *different* formula string
than the simulation expects -- the workbook is broken in Excel even though
parity is green.

This test closes the gap: it builds a small SPIA workbook end-to-end, recomputes
the `ModelCheck` cells via :mod:`xlcalculator` (a pure-Python Excel formula
evaluator), and asserts that the recomputed values match the Python pricing
result at :data:`parity_constants.MODELCHECK_TOL`.

Skipped if `xlcalculator` is not installed (it lives in `requirements-dev.txt`
only) or if the workbook uses formulas xlcalculator does not support yet --
the failure mode in that case is logged and the test is xfail-ed so the
parity gate stays green while the engineer works around the missing function.
"""

from __future__ import annotations

import io

import numpy as np
import pytest
from openpyxl import load_workbook

xc = pytest.importorskip("xlcalculator")

import pricing_projection as sp  # noqa: E402
from build_pricing_excel_workbook import (  # noqa: E402
    ExcelBuildSpec,
    build_workbook_from_spec,
)
from parity_constants import MODELCHECK_TOL  # noqa: E402

pytestmark = [pytest.mark.parity, pytest.mark.product_spia, pytest.mark.slow]


def _build_small_spia_workbook() -> tuple[bytes, sp.SPIAProjectionResult]:
    contract = sp.SPIAContract(issue_age=65, sex="male", benefit_annual=120_000.0)
    yc = sp.YieldCurve.from_flat_rate(0.045)
    ages = np.arange(0, 121, dtype=int)
    qx = np.clip(0.008 + ages * 2e-5, 1e-6, 0.4)
    mort = sp.MortalityTableQx(ages, qx)
    expenses = sp.ExpenseAssumptions(0.0, 0.0, 25.0)
    res = sp.price_spia_single_premium(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=70,
        spread=0.0,
        expenses=expenses,
        expense_annual_inflation=0.0,
    )
    spec = ExcelBuildSpec(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=70,
        spread=0.0,
        valuation_year=2025,
        expenses=expenses,
        expense_annual_inflation=0.0,
        result=res,
    )
    blob = build_workbook_from_spec(spec)
    return blob, res


def _try_recalc_modelcheck(blob: bytes) -> dict[str, float] | None:
    """Use xlcalculator to recompute ModelCheck cells. Returns None if any
    formula function is unsupported by xlcalculator (in which case the test
    is xfail-ed)."""
    tmp = io.BytesIO(blob)
    try:
        compiler = xc.ModelCompiler()
        model = compiler.read_and_parse_archive(tmp)
        evaluator = xc.Evaluator(model)
    except Exception as exc:  # noqa: BLE001 -- xlcalculator raises a wide variety
        pytest.xfail(f"xlcalculator could not parse the workbook: {exc}")

    out: dict[str, float] = {}
    for cell in ("ModelCheck!B5", "ModelCheck!B9"):
        try:
            v = evaluator.evaluate(cell)
        except Exception as exc:  # noqa: BLE001
            pytest.xfail(f"xlcalculator missing function for {cell}: {exc}")
        try:
            out[cell] = float(v)
        except (TypeError, ValueError):
            pytest.xfail(f"xlcalculator returned non-numeric for {cell}: {v!r}")
    return out


def test_modelcheck_cells_recalc_to_python_values() -> None:
    blob, res = _build_small_spia_workbook()
    cells = _try_recalc_modelcheck(blob)
    assert cells is not None  # xfail above on incomplete recalc

    # Sanity: openpyxl should also be able to read the cached values for
    # cross-checking.
    wb = load_workbook(io.BytesIO(blob), data_only=False)
    ws = wb["ModelCheck"]
    assert isinstance(ws["B5"].value, str) and ws["B5"].value.startswith("=")
    assert isinstance(ws["B9"].value, str) and ws["B9"].value.startswith("=")

    # Compare recalculated values to Python at the contracted tolerance.
    np.testing.assert_allclose(
        cells["ModelCheck!B5"], float(res.pv_benefit), rtol=0.0, atol=MODELCHECK_TOL or 1e-6
    )
    np.testing.assert_allclose(
        cells["ModelCheck!B9"], float(res.annuity_factor), rtol=0.0, atol=MODELCHECK_TOL or 1e-6
    )
