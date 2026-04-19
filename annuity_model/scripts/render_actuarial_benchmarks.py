"""render_actuarial_benchmarks -- regenerate the band tables in the docs.

The Section 13.3 benchmark bands and Section 13.4 sensitivity epsilons
listed in :file:`docs/actuarial_benchmarks.md` MUST agree with the
constants exported from :mod:`actuarial_benchmarks`. The doc carries
marker lines of the form::

    <!-- BEGIN GENERATED bands -->
    ...table...
    <!-- END GENERATED bands -->

and this script rewrites the block between them. CI invokes the script
in ``--check`` mode; a drift produces a non-zero exit code.

This mirrors the same pattern as
:mod:`scripts.render_parity_contract`. New products MUST add their
band rows here before commit.

Usage
-----

Regenerate in place::

    python scripts/render_actuarial_benchmarks.py

Verify (used by ``just preflight``)::

    python scripts/render_actuarial_benchmarks.py --check
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

_HERE = Path(__file__).resolve()
_REPO_ROOT = _HERE.parent.parent
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

import actuarial_benchmarks as ab  # noqa: E402

DOCS_DIR = _REPO_ROOT / "docs"
DOC_PATH = DOCS_DIR / "actuarial_benchmarks.md"

BEGIN_MARKER = "<!-- BEGIN GENERATED bands -->"
END_MARKER = "<!-- END GENERATED bands -->"


# (display name, attribute, units). Order matches Section 13.3 of the
# rollout plan.
_BAND_ROWS: tuple[tuple[str, str, str], ...] = (
    # MYGA
    ("MYGA AV(T) lower", "MYGA_BENCHMARK_AV_T_LO", "USD"),
    ("MYGA AV(T) upper", "MYGA_BENCHMARK_AV_T_HI", "USD"),
    ("MYGA PV(maturity) lower", "MYGA_BENCHMARK_PV_LO", "USD"),
    ("MYGA PV(maturity) upper", "MYGA_BENCHMARK_PV_HI", "USD"),
    ("MYGA closed-form AV(T) tolerance", "MYGA_CLOSED_FORM_AV_TOL", "USD"),
    ("MYGA sensitivity epsilon", "MYGA_SENSITIVITY_EPS", "USD"),
    # FIA
    ("FIA AV(T) lower", "FIA_BENCHMARK_AV_T_LO", "USD"),
    ("FIA AV(T) upper", "FIA_BENCHMARK_AV_T_HI", "USD"),
    ("FIA sensitivity epsilon", "FIA_SENSITIVITY_EPS", "USD"),
    # VA
    ("VA AV(T) flat-S&P lower", "VA_BENCHMARK_AV_T_FLAT_LO", "USD"),
    ("VA AV(T) flat-S&P upper", "VA_BENCHMARK_AV_T_FLAT_HI", "USD"),
    ("VA E[AV(T)] MC lower", "VA_BENCHMARK_AV_T_MC_LO", "USD"),
    ("VA E[AV(T)] MC upper", "VA_BENCHMARK_AV_T_MC_HI", "USD"),
    ("VA sensitivity epsilon", "VA_SENSITIVITY_EPS", "USD"),
    # WL
    ("WL single premium lower", "WL_BENCHMARK_SP_LO", "USD"),
    ("WL single premium upper", "WL_BENCHMARK_SP_HI", "USD"),
    ("WL NSP closed-form tolerance", "WL_NSP_TOL", "USD"),
    ("WL sensitivity epsilon", "WL_SENSITIVITY_EPS", "USD"),
    # UL
    ("UL AV(20y) lower", "UL_BENCHMARK_AV_20Y_LO", "USD"),
    ("UL AV(20y) upper", "UL_BENCHMARK_AV_20Y_HI", "USD"),
    ("UL depletion age lower", "UL_BENCHMARK_DEPLETION_AGE_LO", "Years"),
    ("UL depletion age upper", "UL_BENCHMARK_DEPLETION_AGE_HI", "Years"),
    ("UL sensitivity epsilon", "UL_SENSITIVITY_EPS", "USD"),
    # IUL
    ("IUL AV(20y) lower", "IUL_BENCHMARK_AV_20Y_LO", "USD"),
    ("IUL AV(20y) upper", "IUL_BENCHMARK_AV_20Y_HI", "USD"),
    ("IUL sensitivity epsilon", "IUL_SENSITIVITY_EPS", "USD"),
    # VUL
    ("VUL E[AV(20y)] MC lower", "VUL_BENCHMARK_AV_20Y_MC_LO", "USD"),
    ("VUL E[AV(20y)] MC upper", "VUL_BENCHMARK_AV_20Y_MC_HI", "USD"),
    ("VUL sensitivity epsilon", "VUL_SENSITIVITY_EPS", "USD"),
    # Portfolio
    ("Portfolio total CF sum lower", "PORTFOLIO_TOTAL_CF_SUM_LO", "USD"),
    ("Portfolio total CF sum upper", "PORTFOLIO_TOTAL_CF_SUM_HI", "USD"),
    ("Portfolio duration gap lower", "PORTFOLIO_DURATION_GAP_LO", "Years"),
    ("Portfolio duration gap upper", "PORTFOLIO_DURATION_GAP_HI", "Years"),
    ("Portfolio rollup sum consistency tol", "PORTFOLIO_SUM_CONSISTENCY_TOL", "abs"),
)


def _format_value(value: float | int) -> str:
    if isinstance(value, int):
        return f"{value:,}"
    if value == 0.0:
        return "0.0"
    if abs(value) >= 1.0:
        return f"{value:,.4g}"
    return f"{value:.0e}"


def _render_table() -> str:
    lines = [
        "| Quantity | Value | Units | Constant |",
        "|----------|-------|-------|----------|",
    ]
    for display_name, attr, units in _BAND_ROWS:
        value = getattr(ab, attr)
        lines.append(
            f"| {display_name} | `{_format_value(value)}` | {units} "
            f"| `actuarial_benchmarks.{attr}` |"
        )
    # Reflection invariant: every public constant must appear exactly once.
    expected = set(ab.__all__)
    seen = {attr for _, attr, _ in _BAND_ROWS}
    extras = expected - seen
    missing = seen - expected
    if extras or missing:
        raise SystemExit(
            f"render_actuarial_benchmarks: rows out of sync with actuarial_benchmarks.__all__:\n"
            f"  not in __all__: {sorted(missing)!r}\n"
            f"  not in rows:    {sorted(extras)!r}"
        )
    return "\n".join(lines)


def _block() -> str:
    return f"{BEGIN_MARKER}\n{_render_table()}\n{END_MARKER}"


def _replace_block(text: str, new_block: str, *, doc: Path) -> str:
    if BEGIN_MARKER not in text or END_MARKER not in text:
        sys.stderr.write(
            f"{doc}: missing BEGIN/END GENERATED markers; insert these "
            "around the band table to enable rendering.\n"
        )
        raise SystemExit(2)
    pre, _, rest = text.partition(BEGIN_MARKER)
    _, _, post = rest.partition(END_MARKER)
    return f"{pre}{new_block}{post}"


def _process(path: Path, new_block: str, *, check: bool) -> bool:
    if not path.exists():
        sys.stderr.write(f"{path}: missing\n")
        return False
    text = path.read_text(encoding="utf-8")
    rendered = _replace_block(text, new_block, doc=path)
    if rendered == text:
        return True
    if check:
        sys.stderr.write(
            f"{path}: actuarial benchmark table is out of date. Run "
            "`python scripts/render_actuarial_benchmarks.py` to refresh.\n"
        )
        return False
    path.write_text(rendered, encoding="utf-8")
    sys.stdout.write(f"{path}: updated\n")
    return True


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="Verify the doc matches actuarial_benchmarks without writing.",
    )
    args = parser.parse_args(argv)
    return 0 if _process(DOC_PATH, _block(), check=args.check) else 1


if __name__ == "__main__":
    raise SystemExit(main())
