"""Runtime Excel recalc parity gate (P0 hardening, 2026-04 -- restored).

The classic parity tests compare Python to a *Python simulation* of the Excel
formulas (``tests/parity/excel_formula_sim.py``). That guards against logic
drift, but not against a builder bug that emits a *different* formula string
than the simulation expects -- the workbook is broken in Excel even though
parity is green.

This test closes the gap: it builds a small SPIA workbook end-to-end,
recomputes it via LibreOffice headless (``soffice --headless --calc
--convert-to xlsx``), and asserts the cached ``ModelCheck`` cells match the
Python pricing result at :data:`parity_constants.MODELCHECK_TOL`.

History
-------
The original P4 implementation depended on ``xlcalculator``, which became
incompatible with ``numpy>=2`` via its transitive ``yearfrac<2`` pin. The
gate was parked in 2026-04 and the test self-skipped via
``pytest.importorskip("xlcalculator")``. Pure-Python alternatives
(``formulas``, ``pycel``) install cleanly but take >3 minutes to
recalculate the SPIA workbook (full dependency graph build) -- unusable in
CI.

LibreOffice headless is the pragmatic restore: same engine real users open
these workbooks in, ~5 seconds per recalc, installable on every CI runner
and developer laptop. The runtime recalc helper is in
:mod:`excel_runtime_recalc`; it skips gracefully when soffice is not on
PATH so contributors without LibreOffice locally are not blocked.
"""

from __future__ import annotations

import io

import numpy as np
import pytest
from openpyxl import load_workbook

import pricing_projection as sp
from build_pricing_excel_workbook import (
    excel_spec_from_launcher,
    build_workbook_from_spec,
)
from excel_runtime_recalc import (
    LIBREOFFICE_INSTALL_HINT,
    libreoffice_available,
    read_recalculated_cells,
    recalc_workbook,
)
from parity_constants import MODELCHECK_TOL

pytestmark = [pytest.mark.parity, pytest.mark.product_spia, pytest.mark.slow]


def _build_small_spia_workbook() -> tuple[bytes, sp.SPIAProjectionResult]:
    """Build the smallest realistic SPIA workbook so soffice recalc is fast.

    horizon_age=66 gives 12 projection months (one year), which is the
    minimum that exercises the full ModelCheck stack without the heavyweight
    60-month ALM ladder dominating recalc time.
    """
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
        horizon_age=66,
        spread=0.0,
        expenses=expenses,
        expense_annual_inflation=0.0,
    )
    spec = excel_spec_from_launcher(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=66,
        spread=0.0,
        valuation_year=2025,
        expenses=expenses,
        yield_mode_label="flat",
        mortality_mode_label="qx",
        expense_mode_label="manual",
        index_s0=100.0,
        index_levels_at_payment=np.full(12, 100.0),
        expense_annual_inflation=0.0,
    )
    blob = build_workbook_from_spec(spec)
    return blob, res


@pytest.fixture(scope="module")
def libreoffice_or_skip() -> None:
    if not libreoffice_available():
        pytest.skip(
            "LibreOffice (soffice) not on PATH; runtime recalc gate skipped. "
            f"{LIBREOFFICE_INSTALL_HINT}"
        )


def test_modelcheck_cells_recalc_to_python_values(libreoffice_or_skip: None) -> None:
    """ModelCheck!B5 (PV benefit) and ModelCheck!B9 (annuity factor) must
    recalc within Excel's actual formula engine to the Python values within
    ``MODELCHECK_TOL``."""
    blob, res = _build_small_spia_workbook()

    # Sanity: the formula strings are present in the as-built workbook.
    wb = load_workbook(io.BytesIO(blob), data_only=False)
    ws = wb["ModelCheck"]
    assert isinstance(ws["B5"].value, str) and ws["B5"].value.startswith("=")
    assert isinstance(ws["B9"].value, str) and ws["B9"].value.startswith("=")

    recalculated = recalc_workbook(blob, timeout=120.0)
    cells = read_recalculated_cells(
        recalculated, ["ModelCheck!B5", "ModelCheck!B9"]
    )

    assert cells["ModelCheck!B5"] is not None, (
        "soffice did not produce a cached value for ModelCheck!B5; the "
        "workbook may have a recalc-time error (open it in Excel and look "
        "for #VALUE! / #NAME? in the ModelCheck sheet)."
    )
    assert cells["ModelCheck!B9"] is not None, (
        "soffice did not produce a cached value for ModelCheck!B9; see B5 hint."
    )

    np.testing.assert_allclose(
        float(cells["ModelCheck!B5"]),
        float(res.pv_benefit),
        rtol=0.0,
        atol=MODELCHECK_TOL or 1e-6,
        err_msg=(
            "Runtime Excel recalc disagrees with Python on PV benefit. "
            "This means the emitted formula string and the Python pricing "
            "engine compute different values -- check the ModelCheck!B5 "
            "formula and the corresponding python path in pricing_projection."
        ),
    )
    np.testing.assert_allclose(
        float(cells["ModelCheck!B9"]),
        float(res.annuity_factor),
        rtol=0.0,
        atol=MODELCHECK_TOL or 1e-6,
        err_msg=(
            "Runtime Excel recalc disagrees with Python on annuity factor. "
            "Same diagnostic as B5 above."
        ),
    )
