"""parity_trace -- export step-by-step ALM state for Python vs Excel.

A debugging utility that runs the Python ALM engine and the
``excel_formula_sim`` Excel-formula simulation side-by-side for a single
SPIA contract, producing a CSV with one row per projection month and a
``diff_*`` column for every tracked metric.

Use this when an ``investigate_parity_break`` runbook step says "look at
where the trace first diverges". The output CSV can be opened directly
in Excel or pandas; the script also prints the first month at which any
metric exceeds :data:`parity_constants.TOL_DOLLAR`.

Examples
--------
Run with the default scenario (60 month horizon, $250k single premium)::

    python -m annuity_model.scripts.parity_trace --output traces/spia.csv

Restrict to the first 24 months and a tighter dollar threshold::

    python -m annuity_model.scripts.parity_trace --steps 24 --threshold 1e-6

Notes
-----
The script is intentionally narrow: it covers the SPIA + ALM ladder
parity surface only. Other products can implement an analogous trace by
copying this file and replacing :func:`_run_step`.
"""

from __future__ import annotations

import argparse
import csv
import sys
from pathlib import Path

import numpy as np

_HERE = Path(__file__).resolve()
_REPO_ROOT = _HERE.parent.parent
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

import pricing_projection as sp  # noqa: E402
from parity_constants import TOL_DOLLAR  # noqa: E402
from tests.parity.excel_formula_sim import (  # noqa: E402
    excel_disinvest_shortest_first,
)

try:  # noqa: E402
    from excel_runtime_recalc import (
        libreoffice_available,
        read_recalculated_cells,
        recalc_workbook,
    )

    _HAS_RUNTIME_RECALC = True
except ImportError:  # pragma: no cover -- defensive
    _HAS_RUNTIME_RECALC = False


def _default_scenario() -> tuple[
    sp.SPIAContract,
    sp.YieldCurve,
    sp.MortalityTableQx,
    sp.ALMAssumptions,
    sp.ExpenseAssumptions,
]:
    contract = sp.SPIAContract(issue_age=65, sex="male", benefit_annual=120_000.0)
    yc = sp.YieldCurve.from_flat_rate(0.045)
    ages = np.arange(0, 121, dtype=int)
    qx = np.clip(0.008 + ages * 2e-5, 1e-6, 0.4)
    mort = sp.MortalityTableQx(ages, qx)
    alm = sp.ALMAssumptions(
        allocation=sp.alm_default_allocation_spec(),
        rebalance_band=0.10,
        rebalance_frequency_months=1,
        reinvest_rule="hold_cash",
        disinvest_rule="shortest_first",
        liquidity_near_liquid_years=0.25,
    )
    expenses = sp.ExpenseAssumptions(0.0, 0.0, 25.0)
    return contract, yc, mort, alm, expenses


def _row_for_month(
    month: int, py_state: dict[str, float], xl_state: dict[str, float]
) -> dict[str, object]:
    """Build a flat dict row with python, excel, and diff columns."""
    row: dict[str, object] = {"month": month}
    for key, val in py_state.items():
        row[f"py_{key}"] = val
    for key, val in xl_state.items():
        row[f"xl_{key}"] = val
    for key in py_state:
        try:
            row[f"diff_{key}"] = float(py_state[key]) - float(xl_state.get(key, float("nan")))
        except (TypeError, ValueError):
            row[f"diff_{key}"] = "n/a"
    return row


def _excel_states_via_libreoffice(
    *,
    contract: sp.SPIAContract,
    yc: sp.YieldCurve,
    mort: sp.MortalityTableQx,
    alm: sp.ALMAssumptions,
    expenses: sp.ExpenseAssumptions,
    pricing: sp.SPIAProjectionResult,
    py_alm: sp.ALMResult,
    horizon_age: int,
    n_months: int,
) -> list[dict[str, float]] | None:
    """Build a SPIA + ALM workbook, recalc via LibreOffice, and read the
    cached per-month ALM_Projection cells. Returns one dict per month with
    asset_mv / liab_pv / surplus / borrowing pulled from the cached cells.

    Returns ``None`` if LibreOffice is not available so callers can fall
    back to a clearly-marked NaN trace instead of silently mirroring Python.
    """
    if not _HAS_RUNTIME_RECALC or not libreoffice_available():
        return None

    # Local imports to keep parity_trace usable when build_pricing_excel_workbook
    # transitively breaks (e.g. mid-refactor); the trace falls back gracefully
    # rather than refusing to run.
    from build_pricing_excel_workbook import (
        ALM_PROJECTION_FIRST_DATA_ROW,
        alm_excel_snapshot_from_result,
        build_workbook_from_spec,
        excel_spec_from_launcher,
    )

    spec = excel_spec_from_launcher(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=horizon_age,
        spread=0.0,
        valuation_year=2025,
        expenses=expenses,
        yield_mode_label="flat",
        mortality_mode_label="qx",
        expense_mode_label="manual",
        index_s0=100.0,
        index_levels_at_payment=np.full(int(pricing.times_years.shape[0]), 100.0),
        expense_annual_inflation=0.0,
    )
    alm_snap = alm_excel_snapshot_from_result(py_alm, alm)
    blob = build_workbook_from_spec(
        spec, alm_snapshot=alm_snap, alm_assumptions=alm
    )

    recalculated = recalc_workbook(blob, timeout=180.0)
    addrs: list[str] = []
    # Columns on ALM_Projection: C=Asset MV, D=Liability PV, E=Borrowing,
    # F=Surplus. First data row = ALM_PROJECTION_FIRST_DATA_ROW (=13).
    for i in range(n_months):
        r = ALM_PROJECTION_FIRST_DATA_ROW + i
        addrs.extend(
            [
                f"ALM_Projection!C{r}",
                f"ALM_Projection!D{r}",
                f"ALM_Projection!E{r}",
                f"ALM_Projection!F{r}",
            ]
        )
    cells = read_recalculated_cells(recalculated, addrs)

    out: list[dict[str, float]] = []
    nan = float("nan")
    for i in range(n_months):
        r = ALM_PROJECTION_FIRST_DATA_ROW + i

        def _f(addr: str) -> float:
            v = cells.get(addr)
            try:
                return float(v) if v is not None else nan
            except (TypeError, ValueError):
                return nan

        out.append(
            {
                "asset_mv": _f(f"ALM_Projection!C{r}"),
                "liab_pv": _f(f"ALM_Projection!D{r}"),
                "borrowing": _f(f"ALM_Projection!E{r}"),
                "surplus": _f(f"ALM_Projection!F{r}"),
            }
        )
    return out


def _trace(horizon_months: int) -> list[dict[str, object]]:
    contract, yc, mort, alm, expenses = _default_scenario()

    horizon_age = contract.issue_age + (horizon_months // 12) + 1
    pricing = sp.price_spia_single_premium(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=horizon_age,
        spread=0.0,
        expenses=expenses,
        expense_annual_inflation=0.0,
    )
    py = sp.run_alm_projection_from_pricing_result(
        pricing=pricing, yield_curve=yc, spread=0.0, assumptions=alm
    )

    n = min(horizon_months, py.surplus.shape[0])

    excel_states: list[dict[str, float]] | None = None
    try:
        excel_states = _excel_states_via_libreoffice(
            contract=contract,
            yc=yc,
            mort=mort,
            alm=alm,
            expenses=expenses,
            pricing=pricing,
            py_alm=py,
            horizon_age=horizon_age,
            n_months=n,
        )
    except Exception as exc:  # noqa: BLE001 -- soffice is unreliable on dev laptops
        sys.stderr.write(
            f"warning: LibreOffice recalc failed ({exc!r}); "
            f"Excel-side columns will be NaN for this trace.\n"
        )

    if excel_states is None:
        sys.stderr.write(
            "warning: LibreOffice not available; Excel-side trace columns "
            "are NaN. Install libreoffice-calc to enable real Excel recalc.\n"
        )

    rows: list[dict[str, object]] = []
    nan = float("nan")
    for m in range(n):
        py_state = {
            "asset_mv": float(py.asset_market_value[m]),
            "liab_pv": float(py.liability_pv[m]),
            "surplus": float(py.surplus[m]),
            "borrowing": float(py.borrowing_balance[m]),
        }
        if excel_states is not None and m < len(excel_states):
            xl_state = {
                "asset_mv": excel_states[m]["asset_mv"],
                "liab_pv": excel_states[m]["liab_pv"],
                "surplus": excel_states[m]["surplus"],
                "borrowing": excel_states[m]["borrowing"],
            }
        else:
            # Sentinel NaN -- callers see this in the diff_* columns and know
            # the Excel side did not contribute. Critically we DO NOT mirror
            # py_state, so a green trace cannot accidentally hide drift.
            xl_state = {k: nan for k in py_state}
        rows.append(_row_for_month(m + 1, py_state, xl_state))
    return rows


def export_trace(output_path: Path, *, steps: int, threshold: float) -> int:
    """Write the trace CSV. Return 0 on success, 1 if any metric breaches ``threshold``."""
    rows = _trace(steps)
    if not rows:
        sys.stderr.write("trace produced no rows\n")
        return 2

    output_path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = list(rows[0].keys())
    with output_path.open("w", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    sys.stdout.write(f"Trace written to {output_path} ({len(rows)} rows)\n")

    import math

    breaches: list[tuple[int, str, float]] = []
    nan_diffs: list[tuple[int, str]] = []
    for row in rows:
        for key, val in row.items():
            if not key.startswith("diff_"):
                continue
            if not isinstance(val, float):
                continue
            if math.isnan(val):
                # NaN diff means the Excel side could not contribute -- record
                # so the user sees that the trace is one-sided rather than
                # treating "abs(nan) > threshold == False" as a pass.
                nan_diffs.append((int(row["month"]), key))  # type: ignore[arg-type]
                continue
            if abs(val) > threshold:
                breaches.append((int(row["month"]), key, val))  # type: ignore[arg-type]

    if breaches:
        sys.stdout.write(f"\nBreaches (|diff| > {threshold}):\n")
        for month, key, val in breaches[:25]:
            sys.stdout.write(f"  month {month:>4d}  {key:<20s} = {val:+.6e}\n")
        if len(breaches) > 25:
            sys.stdout.write(f"  ... {len(breaches) - 25} more\n")
        return 1

    if nan_diffs:
        sys.stdout.write(
            f"All numeric diffs within tolerance ({threshold}), but "
            f"{len(nan_diffs)} entries are NaN (Excel side missing). "
            f"Install libreoffice-calc to populate the Excel side.\n"
        )
        return 0

    sys.stdout.write(f"All metrics within tolerance ({threshold}).\n")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("traces/parity_trace.csv"),
        help="Output CSV path (default: %(default)s)",
    )
    parser.add_argument(
        "--steps",
        type=int,
        default=60,
        help="Number of projection months to trace (default: %(default)s)",
    )
    parser.add_argument(
        "--threshold",
        type=float,
        default=TOL_DOLLAR,
        help="|diff| threshold for breach reporting (default: parity_constants.TOL_DOLLAR)",
    )
    args = parser.parse_args(argv)
    return export_trace(args.output, steps=args.steps, threshold=args.threshold)


# Surface the helper so the runbook can import it directly.
__all__ = ["export_trace", "main"]
_ = excel_disinvest_shortest_first  # keep import to make available for adapters

if __name__ == "__main__":
    raise SystemExit(main())
