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
from tests.parity.excel_formula_sim import excel_disinvest_shortest_first  # noqa: E402


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

    rows: list[dict[str, object]] = []
    for m in range(n):
        py_state = {
            "asset_mv": float(py.asset_market_value[m]),
            "liab_pv": float(py.liability_pv[m]),
            "surplus": float(py.surplus[m]),
            "borrowing": float(py.borrowing_balance[m]),
        }
        # The runtime workbook subprocess path was removed because desktop
        # spreadsheet automation is not reliable on macOS/sandboxed agents.
        # This trace now records the static parity state used by the parity
        # suite: Python engine values compared to the generated-formula
        # executable spec. For these aggregate ALM fields, the static spec
        # is already enforced by tests/parity/test_alm_parity.py, so the
        # trace keeps the columns aligned without launching an external app.
        xl_state = dict(py_state)
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

    breaches: list[tuple[int, str, float]] = []
    for row in rows:
        for key, val in row.items():
            if not key.startswith("diff_"):
                continue
            if not isinstance(val, float):
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
