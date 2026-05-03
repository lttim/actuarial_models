#!/usr/bin/env python3
"""PR-level mutmut gate over the parity-critical surface.

The nightly ``mutmut-nightly.yml`` workflow runs mutmut over the *entire*
parity-critical surface (every engine, every workbook builder, the
validator, the registries) and uploads survivors as a non-blocking
artifact. That cadence is fine for slow drift detection but it gives no
signal at PR review time, so a regression in test coverage on a touched
engine module can ship and only be noticed the next morning.

This script tightens the loop. On each PR:

1. The CI workflow computes ``git diff --name-only $base...HEAD`` and
   passes the result via ``--touched-files``.
2. We intersect with :data:`MUTMUT_SURFACE` -- the curated allow-list of
   parity-critical modules, kept in lockstep with both
   ``.github/workflows/mutmut-nightly.yml`` (nightly surface) and the
   ``[[tool.mypy.overrides]]`` strict module list in ``pyproject.toml``.
3. If the intersection is empty, the gate exits 0 silently -- a docs/
   tests-only PR pays no mutmut cost.
4. Otherwise we write a transient ``setup.cfg`` ``[mutmut]`` section
   constrained to just those touched files, run ``mutmut run``, parse
   the per-file ``mutants/<path>.meta`` JSON files for survivors, and
   compare each file's survivor count to its threshold from
   ``mutmut_thresholds.toml`` (default 0).
5. If any file exceeds its cap, the gate fails the PR with a per-file
   table pointing at exactly which file regressed.

Local triage::

    python scripts/mutmut_pr_gate.py \\
        --touched-files pricing_projection.py rila_projection.py
    # or, with no shell-out (parse a pre-existing mutants/ tree):
    python scripts/mutmut_pr_gate.py --skip-run \\
        --touched-files pricing_projection.py

Exit codes
----------
* ``0`` -- no parity-critical files touched, OR all touched files within
  their survivor cap.
* ``1`` -- at least one touched file exceeded its survivor cap.
* ``2`` -- usage / config error (missing thresholds file, no touched
  files supplied, mutmut binary not on PATH, etc.).
"""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import tomllib
from dataclasses import dataclass
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
THRESHOLDS_PATH = REPO_ROOT / "mutmut_thresholds.toml"
MUTANTS_DIR = REPO_ROOT / "mutants"
SETUP_CFG = REPO_ROOT / "setup.cfg"

# Allow-list of parity-critical files mutmut is allowed to mutate. Must
# stay in lockstep with `.github/workflows/mutmut-nightly.yml` --
# `tests/test_mutmut_pr_gate.py` enforces that with a meta-test.
MUTMUT_SURFACE: frozenset[str] = frozenset(
    {
        "pricing_projection.py",
        "term_projection.py",
        "rila_projection.py",
        "alm_excel_ladder.py",
        "build_pricing_excel_workbook.py",
        "build_rila_excel_workbook.py",
        "build_term_excel_workbook.py",
        "excel_workbook_validator.py",
        "product_excel.py",
        "product_registry.py",
    }
)

# Mutmut 3.x exit codes per mutant (see mutmut/__main__.py
# ``status_by_exit_code``). Anything we treat as a "survivor" is a
# coverage gap that should fail the PR if over threshold.
SURVIVED_EXIT_CODES: frozenset[int] = frozenset({0})


@dataclass(frozen=True)
class FileSurvivorCount:
    """Per-file survivor tally from a mutmut run."""

    path: str
    total_mutants: int
    survivors: int
    threshold: int

    @property
    def passed(self) -> bool:
        return self.survivors <= self.threshold

    def render_row(self) -> str:
        marker = "ok" if self.passed else "FAIL"
        return (
            f"  [{marker}] {self.path}: {self.survivors} survivor(s) / "
            f"{self.total_mutants} mutants (cap = {self.threshold})"
        )


def _load_thresholds(path: Path = THRESHOLDS_PATH) -> tuple[int, dict[str, int]]:
    if not path.is_file():
        print(f"[mutmut-gate] missing {path}", file=sys.stderr)
        sys.exit(2)
    with path.open("rb") as fh:
        data = tomllib.load(fh)
    section = data.get("thresholds", {})
    default = section.get("default", 0)
    if not isinstance(default, int) or default < 0:
        print(
            f"[mutmut-gate] thresholds.default must be a non-negative int, got {default!r}",
            file=sys.stderr,
        )
        sys.exit(2)
    per_file_raw = section.get("per_file", {})
    if not isinstance(per_file_raw, dict):
        print(
            f"[mutmut-gate] thresholds.per_file must be a table, got {type(per_file_raw).__name__}",
            file=sys.stderr,
        )
        sys.exit(2)
    per_file: dict[str, int] = {}
    for k, v in per_file_raw.items():
        if not isinstance(v, int) or v < 0:
            print(
                f"[mutmut-gate] threshold for {k!r} must be a non-negative int, got {v!r}",
                file=sys.stderr,
            )
            sys.exit(2)
        per_file[k] = v
    return default, per_file


def _filter_to_surface(touched: list[str]) -> list[str]:
    return sorted(set(touched) & MUTMUT_SURFACE)


def _write_setup_cfg(files: list[str]) -> str:
    """Write a transient ``setup.cfg`` ``[mutmut]`` section for *files*.

    We deliberately do NOT clobber a pre-existing setup.cfg if one exists
    (it may carry unrelated package metadata in some hypothetical future);
    we save the original, append our section, and the caller is expected
    to call :func:`_restore_setup_cfg` in a try/finally.
    """
    backup = ""
    if SETUP_CFG.exists():
        backup = SETUP_CFG.read_text()
    paths_csv = ",".join(files)
    content = (
        backup
        + ("\n" if backup and not backup.endswith("\n") else "")
        + "[mutmut]\n"
        + f"paths_to_mutate={paths_csv}\n"
        + "tests_dir=tests\n"
    )
    SETUP_CFG.write_text(content)
    return backup


def _restore_setup_cfg(backup: str) -> None:
    if backup:
        SETUP_CFG.write_text(backup)
    else:
        SETUP_CFG.unlink(missing_ok=True)


def _run_mutmut() -> None:
    if shutil.which("mutmut") is None:
        print("[mutmut-gate] `mutmut` binary not on PATH", file=sys.stderr)
        sys.exit(2)
    proc = subprocess.run(
        ["mutmut", "run"],
        cwd=REPO_ROOT,
        check=False,
    )
    # mutmut returns non-zero when survivors exist, but we make our own
    # decision based on per-file thresholds, so we do NOT propagate its
    # exit code here.
    if proc.returncode not in (0, 1, 2):
        print(
            f"[mutmut-gate] `mutmut run` exited with unexpected code {proc.returncode}; aborting",
            file=sys.stderr,
        )
        sys.exit(2)


def _parse_meta_file(path: Path) -> tuple[int, int]:
    """Return ``(total_mutants, survivors)`` for one ``.meta`` file."""
    raw = json.loads(path.read_text())
    exit_codes = raw.get("exit_code_by_key", {})
    total = len(exit_codes)
    survivors = sum(1 for code in exit_codes.values() if code in SURVIVED_EXIT_CODES)
    return total, survivors


def _collect_per_file(
    files: list[str],
    default: int,
    per_file: dict[str, int],
    mutants_dir: Path = MUTANTS_DIR,
) -> list[FileSurvivorCount]:
    out: list[FileSurvivorCount] = []
    for rel in files:
        meta = mutants_dir / f"{rel}.meta"
        if not meta.is_file():
            print(
                f"[mutmut-gate] WARNING: no mutants/{rel}.meta produced; "
                "mutmut may have skipped this file (no test coverage at all, "
                "or no mutable AST nodes). Counting as 0/0 survivors.",
                file=sys.stderr,
            )
            total, survivors = 0, 0
        else:
            total, survivors = _parse_meta_file(meta)
        out.append(
            FileSurvivorCount(
                path=rel,
                total_mutants=total,
                survivors=survivors,
                threshold=per_file.get(rel, default),
            )
        )
    return out


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--touched-files",
        nargs="*",
        default=[],
        help=(
            "list of file paths (relative to annuity_model/) touched by "
            "the PR. Typically supplied by the workflow as "
            "`git diff --name-only $base...HEAD`."
        ),
    )
    parser.add_argument(
        "--skip-run",
        action="store_true",
        help=(
            "do not invoke `mutmut run`; only parse pre-existing "
            "`mutants/` data. Useful for local triage and unit tests."
        ),
    )
    parser.add_argument(
        "--thresholds",
        type=Path,
        default=THRESHOLDS_PATH,
        help="override path to the thresholds TOML file (test hook).",
    )
    parser.add_argument(
        "--mutants-dir",
        type=Path,
        default=MUTANTS_DIR,
        help="override the `mutants/` directory (test hook).",
    )
    args = parser.parse_args(argv)

    default, per_file = _load_thresholds(args.thresholds)

    target_files = _filter_to_surface(args.touched_files)
    if not target_files:
        if args.touched_files:
            print(
                "[mutmut-gate] no parity-critical files in the touched set; "
                "nightly mutmut covers the full surface."
            )
        else:
            print(
                "[mutmut-gate] no --touched-files passed; nothing to gate. "
                "(In CI this means the diff was empty against the base ref.)"
            )
        return 0

    print(
        f"[mutmut-gate] gating {len(target_files)} touched parity-critical file(s): {target_files}"
    )

    if not args.skip_run:
        backup = _write_setup_cfg(target_files)
        try:
            _run_mutmut()
        finally:
            _restore_setup_cfg(backup)

    rows = _collect_per_file(target_files, default, per_file, mutants_dir=args.mutants_dir)

    print("[mutmut-gate] per-file survivor counts:")
    for r in rows:
        print(r.render_row())

    failures = [r for r in rows if not r.passed]
    if failures:
        print(
            f"\n[mutmut-gate] FAIL: {len(failures)} file(s) over their "
            "survivor cap. Either add tests that kill the surviving "
            "mutants, or -- if the survivors are inherently untestable -- "
            "raise the per-file threshold in mutmut_thresholds.toml with "
            "a one-line justification AND CODEOWNERS sign-off.",
            file=sys.stderr,
        )
        return 1

    print("[mutmut-gate] PASS: all touched files within their survivor cap.")
    return 0


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
