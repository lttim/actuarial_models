"""
export_parity_trace.py
----------------------
Debugging utility: export step-by-step ALM state from both the Python engine
and the Excel formula simulation to a CSV, so discrepancies can be identified.

Usage (from project root):
    python scripts/export_parity_trace.py --output traces/parity_trace.csv

Adapt the _run_python_step() and _run_excel_step() functions to your product.
"""

from __future__ import annotations

import argparse
import csv
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))


def _run_python_step(state):
    """ADAPT: Run one step of your Python engine. Return (new_state, metrics_dict)."""
    raise NotImplementedError("Implement for your product")


def _run_excel_step(state):
    """ADAPT: Run one step of your Excel formula simulation. Return (new_state, metrics_dict)."""
    raise NotImplementedError("Implement for your product")


def _initial_state():
    """ADAPT: Return the initial state for the projection."""
    raise NotImplementedError("Implement for your product")


def export_trace(output_path: Path, n_steps: int = 60) -> None:
    state_py = _initial_state()
    state_xl = _initial_state()

    rows = []
    for step in range(1, n_steps + 1):
        state_py, metrics_py = _run_python_step(state_py)
        state_xl, metrics_xl = _run_excel_step(state_xl)

        row = {"step": step}
        for key, val in metrics_py.items():
            row[f"py_{key}"] = val
        for key, val in metrics_xl.items():
            row[f"xl_{key}"] = val
        for key in metrics_py:
            py_val = metrics_py.get(key, float("nan"))
            xl_val = metrics_xl.get(key, float("nan"))
            try:
                row[f"diff_{key}"] = py_val - xl_val
            except TypeError:
                row[f"diff_{key}"] = "n/a"
        rows.append(row)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = list(rows[0].keys()) if rows else []
    with open(output_path, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    print(f"Trace exported to: {output_path}")

    # Print summary: steps with non-trivial discrepancy
    print("\nSteps with |diff| > 0.01 in any metric:")
    found = False
    for row in rows:
        diffs = {k: v for k, v in row.items() if k.startswith("diff_") and isinstance(v, float) and abs(v) > 0.01}
        if diffs:
            print(f"  Step {row['step']}: {diffs}")
            found = True
    if not found:
        print("  None — all steps within tolerance.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Export parity trace to CSV")
    parser.add_argument("--output", default="traces/parity_trace.csv", help="Output CSV path")
    parser.add_argument("--steps", type=int, default=60, help="Number of projection steps")
    args = parser.parse_args()
    export_trace(Path(args.output), args.steps)
