#!/usr/bin/env python3
"""One-way coverage ratchet driven by ``pyproject.toml``.

Background
----------
The historical CI invocation hard-coded ``coverage report --fail-under=55``
in ``.github/workflows/ci.yml``. That created two failure modes:

1. Two sources of truth for the gate -- the workflow YAML and the
   ``[tool.coverage.report]`` table in ``pyproject.toml`` -- which silently
   drifted apart.
2. The floor was never raised when test coverage genuinely improved, because
   bumping it required editing CI YAML, which most contributors avoid.

This script collapses both into a single source of truth: the
``fail_under`` key in ``[tool.coverage.report]`` of
``annuity_model/pyproject.toml``. CI calls
``python scripts/ratchet_coverage.py``; the script reads the floor, asks
``coverage`` for the actual percentage, and fails if actual < floor.

Local contributors who improve coverage can run::

    python scripts/ratchet_coverage.py --update

to rewrite the ``fail_under`` value to match today's actual coverage,
producing a clean diff that goes through normal CODEOWNERS review.

The script will *never* lower the floor automatically. Doing so requires
hand-editing ``pyproject.toml`` and is intentionally awkward.

Exit codes
----------
* ``0`` -- coverage >= floor (gate passes).
* ``1`` -- coverage < floor (gate fails).
* ``2`` -- usage error / could not invoke ``coverage`` / pyproject malformed.
* ``3`` -- ``--update`` requested but the ratchet would *lower* the floor;
  refused.
"""

from __future__ import annotations

import argparse
import subprocess
import sys
import tomllib
from pathlib import Path

PYPROJECT_PATH = Path(__file__).resolve().parents[1] / "pyproject.toml"
DEFAULT_BUMP_HINT = 1.0


def _read_floor() -> float:
    """Return the current ratchet floor from ``pyproject.toml``."""
    if not PYPROJECT_PATH.is_file():
        print(f"[ratchet] missing {PYPROJECT_PATH}", file=sys.stderr)
        sys.exit(2)
    with PYPROJECT_PATH.open("rb") as fh:
        data = tomllib.load(fh)
    try:
        floor = data["tool"]["coverage"]["report"]["fail_under"]
    except KeyError:
        print(
            "[ratchet] [tool.coverage.report].fail_under is missing from "
            f"{PYPROJECT_PATH}; cannot enforce ratchet.",
            file=sys.stderr,
        )
        sys.exit(2)
    if not isinstance(floor, int | float):
        print(
            f"[ratchet] fail_under must be a number, got {floor!r} ({type(floor).__name__})",
            file=sys.stderr,
        )
        sys.exit(2)
    return float(floor)


def _measure_actual(coverage_cmd: list[str]) -> float:
    """Return today's coverage percentage by shelling out to ``coverage``.

    We use ``--format=total`` (added in coverage 7.x) which prints exactly
    one number to stdout: the overall percentage. We also pass
    ``--fail-under=0`` so coverage's own gate cannot fire here -- the
    ratchet's whole job is to be the gate, and we need a clean reading no
    matter how low coverage drops.
    """
    try:
        proc = subprocess.run(
            [*coverage_cmd, "report", "--format=total", "--fail-under=0"],
            check=False,
            capture_output=True,
            text=True,
        )
    except FileNotFoundError as exc:
        print(f"[ratchet] could not invoke coverage: {exc}", file=sys.stderr)
        sys.exit(2)
    raw = proc.stdout.strip()
    if proc.returncode != 0 and not raw:
        print(
            "[ratchet] `coverage report --format=total --fail-under=0` "
            f"exited {proc.returncode} with empty stdout; stderr:\n"
            f"{proc.stderr}",
            file=sys.stderr,
        )
        sys.exit(2)
    try:
        return float(raw)
    except ValueError:
        print(
            f"[ratchet] could not parse coverage output as float: {raw!r}; "
            f"stderr:\n{proc.stderr}",
            file=sys.stderr,
        )
        sys.exit(2)


def _rewrite_floor(new_floor: float) -> None:
    """Replace the ``fail_under = X`` line in ``pyproject.toml`` in place.

    We do a textual rewrite (rather than reformatting via tomli-w / tomlkit)
    to avoid disturbing surrounding comments, which carry the rationale and
    the history note. The file is small enough that this is robust.
    """
    text = PYPROJECT_PATH.read_text()
    new_line = f"fail_under = {new_floor:.1f}"
    matches = [
        line for line in text.splitlines() if line.startswith("fail_under")
    ]
    if len(matches) != 1:
        print(
            f"[ratchet] expected exactly one `fail_under = ...` line in "
            f"{PYPROJECT_PATH}, found {len(matches)}: {matches!r}",
            file=sys.stderr,
        )
        sys.exit(2)
    old_line = matches[0]
    PYPROJECT_PATH.write_text(text.replace(old_line, new_line, 1))
    print(f"[ratchet] updated {PYPROJECT_PATH}: '{old_line}' -> '{new_line}'")


def _format_report(actual: float, floor: float) -> str:
    delta = actual - floor
    headroom = "above" if delta >= 0 else "below"
    return (
        f"[ratchet] actual coverage = {actual:.1f}%; "
        f"floor = {floor:.1f}%; {abs(delta):.1f} pp {headroom} floor"
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--update",
        action="store_true",
        help=(
            "rewrite pyproject's fail_under to today's actual coverage. "
            "Refuses to lower the floor; this is the *upward* ratchet only."
        ),
    )
    parser.add_argument(
        "--bump-hint",
        type=float,
        default=DEFAULT_BUMP_HINT,
        help=(
            "print a 'consider bumping' hint when actual exceeds floor by "
            f"this many percentage points (default: {DEFAULT_BUMP_HINT})."
        ),
    )
    parser.add_argument(
        "--coverage-cmd",
        nargs="+",
        default=["coverage"],
        help=(
            "command (and optional args) to invoke `coverage`. Override for "
            "venv-aware CI runners, e.g. `--coverage-cmd python -m coverage`."
        ),
    )
    args = parser.parse_args(argv)

    floor = _read_floor()
    actual = _measure_actual(args.coverage_cmd)
    print(_format_report(actual, floor), flush=True)

    if args.update:
        if actual < floor:
            print(
                f"[ratchet] refusing to update: actual {actual:.1f}% would "
                f"*lower* floor {floor:.1f}%. Manually edit pyproject.toml "
                "and document why in the same PR.",
                file=sys.stderr,
            )
            return 3
        new_floor = round(actual, 1)
        if new_floor <= floor:
            print(
                f"[ratchet] no update needed; rounded actual ({new_floor:.1f}%)"
                f" is not above current floor ({floor:.1f}%)."
            )
            return 0
        _rewrite_floor(new_floor)
        return 0

    if actual + 1e-9 < floor:
        print(
            f"[ratchet] FAIL: coverage {actual:.1f}% < floor {floor:.1f}%. "
            "Either add tests or, if the floor is genuinely too aggressive, "
            "lower it manually in pyproject.toml with reviewer sign-off "
            "and a justification in the PR description.",
            file=sys.stderr,
        )
        return 1

    if actual >= floor + args.bump_hint:
        print(
            f"[ratchet] hint: actual is {actual - floor:.1f} pp above the "
            f"floor; consider running `python scripts/ratchet_coverage.py "
            "--update` in a follow-up PR to lock in the gain."
        )
    return 0


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
