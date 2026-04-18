"""render_parity_contract -- regenerate the tolerance tables in the contracts.

The tolerance tables in :file:`docs/model_parity_contract.md` and
:file:`docs/rila_parity_contract.md` MUST agree with the constants exported
from :mod:`parity_constants`. Rather than maintain that agreement by hand,
each contract document carries marker lines of the form::

    <!-- BEGIN GENERATED tolerances -->
    ...table...
    <!-- END GENERATED tolerances -->

and this script rewrites the block between them. CI invokes the script in
``--check`` mode; a drift produces a non-zero exit code.

Usage
-----

Regenerate in place::

    python -m annuity_model.scripts.render_parity_contract

Verify (used by CI / pre-commit)::

    python -m annuity_model.scripts.render_parity_contract --check
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

_HERE = Path(__file__).resolve()
_REPO_ROOT = _HERE.parent.parent
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

import parity_constants as pc  # noqa: E402

DOCS_DIR = _REPO_ROOT / "docs"

BEGIN_MARKER = "<!-- BEGIN GENERATED tolerances -->"
END_MARKER = "<!-- END GENERATED tolerances -->"


# Each entry: (display name, attribute on parity_constants, units, notes).
_SPIA_ROWS: tuple[tuple[str, str, str, str], ...] = (
    ("Cash balance EOM", "TOL_DOLLAR", "USD", "Per month"),
    ("Bond face EOM", "TOL_DOLLAR", "USD", "Per bucket per month"),
    ("Bond MV EOM", "TOL_DOLLAR", "USD", "Per bucket per month"),
    ("Total asset MV", "TOL_DOLLAR", "USD", ""),
    ("Liability PV", "TOL_DOLLAR", "USD", ""),
    ("Surplus", "TOL_DOLLAR", "USD", ""),
    ("Remaining tenor (t_rem)", "TOL_TENOR", "Years", ""),
    ("Discount factor (DF)", "TOL_DF", "Dimensionless", ""),
    ("Reinvest amount (dmv)", "TOL_DOLLAR", "USD", ""),
    ("ModelCheck snapshot", "MODELCHECK_TOL", "USD", "Exact match required"),
    ("Disinvest tie-break threshold (Excel)", "EXCEL_DISINVEST_THRESHOLD", "Dimensionless", "Half the inter-bucket epsilon"),
    ("Per-bucket epsilon (Python / Excel)", "EXCEL_DISINVEST_EPSILON", "Dimensionless", "k * eps (Py) / (k+1) * eps (Excel)"),
)

_RILA_ROWS: tuple[tuple[str, str, str, str], ...] = (
    ("PV(benefit) cell", "RILA_PV_TOL", "USD", "ModelCheck B5"),
    ("PV(expenses) cell", "RILA_PV_TOL", "USD", "ModelCheck B6"),
    ("PV(total) cell", "RILA_PV_TOL", "USD", "ModelCheck B7"),
    ("Single premium cell", "RILA_PV_TOL", "USD", "ModelCheck B8"),
    ("Account value path", "RILA_AV_TOL", "USD", "Per month"),
    ("ModelCheck snapshot", "MODELCHECK_TOL", "USD", "Exact match required"),
)


def _format_tolerance(value: float) -> str:
    """Format a constant value the way it would appear in the docs."""
    if value == 0.0:
        return "0.0 (exact)"
    if value >= 1e-3:
        return f"{value:.6g}"
    return f"{value:.0e}"


def _render_table(rows: tuple[tuple[str, str, str, str], ...]) -> str:
    lines = [
        "| Variable | Tolerance | Units | Notes |",
        "|----------|-----------|-------|-------|",
    ]
    for display_name, attr, units, notes in rows:
        value = getattr(pc, attr)
        lines.append(
            f"| {display_name} | `{_format_tolerance(value)}` "
            f"(`parity_constants.{attr}`) | {units} | {notes} |"
        )
    return "\n".join(lines)


def _spia_block() -> str:
    return f"{BEGIN_MARKER}\n{_render_table(_SPIA_ROWS)}\n{END_MARKER}"


def _rila_block() -> str:
    return f"{BEGIN_MARKER}\n{_render_table(_RILA_ROWS)}\n{END_MARKER}"


def _replace_block(text: str, new_block: str, *, doc: Path) -> str:
    if BEGIN_MARKER not in text or END_MARKER not in text:
        sys.stderr.write(
            f"{doc}: missing BEGIN/END GENERATED markers; insert these around the "
            "tolerance table to enable rendering.\n"
        )
        raise SystemExit(2)
    pre, _, rest = text.partition(BEGIN_MARKER)
    _, _, post = rest.partition(END_MARKER)
    return f"{pre}{new_block}{post}"


def _process(path: Path, new_block: str, *, check: bool) -> bool:
    """Return True if the file is up to date (or was updated)."""
    if not path.exists():
        sys.stderr.write(f"{path}: missing\n")
        return False
    text = path.read_text(encoding="utf-8")
    rendered = _replace_block(text, new_block, doc=path)
    if rendered == text:
        return True
    if check:
        sys.stderr.write(
            f"{path}: tolerance table is out of date. Run "
            "`python -m annuity_model.scripts.render_parity_contract` to refresh.\n"
        )
        return False
    path.write_text(rendered, encoding="utf-8")
    sys.stdout.write(f"{path}: updated\n")
    return True


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument(
        "--check",
        action="store_true",
        help="Verify the docs match parity_constants without writing.",
    )
    args = parser.parse_args(argv)

    spia_doc = DOCS_DIR / "model_parity_contract.md"
    rila_doc = DOCS_DIR / "rila_parity_contract.md"
    ok = True
    ok &= _process(spia_doc, _spia_block(), check=args.check)
    ok &= _process(rila_doc, _rila_block(), check=args.check)
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
