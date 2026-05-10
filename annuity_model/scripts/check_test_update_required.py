"""Guardrail: behavior changes must carry test changes.

Local usage (staged changes):
    python annuity_model/scripts/check_test_update_required.py

CI usage (PR diff):
    python annuity_model/scripts/check_test_update_required.py <base_sha> <head_sha>

The guard fails when behavior-impacting Python files change without any updated
pytest files under `annuity_model/tests/`.
"""

from __future__ import annotations

import fnmatch

# Reviewed: PR guard invokes fixed local git command vectors with shell=False.
import subprocess  # nosec B404
import sys
from dataclasses import dataclass
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]

BEHAVIOR_GLOBS = (
    "annuity_model/src/annuity_model/*_projection.py",
    "annuity_model/src/annuity_model/build_*_excel_workbook.py",
    "annuity_model/src/annuity_model/alm_excel_ladder.py",
    "annuity_model/src/annuity_model/excel_workbook_validator.py",
    "annuity_model/src/annuity_model/pricing_ui.py",
    "annuity_model/src/annuity_model/cli.py",
    "annuity_model/src/annuity_model/product_*.py",
    "annuity_model/src/annuity_model/liability_*.py",
    "annuity_model/src/annuity_model/portfolio_*.py",
    "annuity_model/src/annuity_model/inforce_*.py",
    "annuity_model/src/annuity_model/pricing_scenario_materialize.py",
    "annuity_model/src/annuity_model/account_value.py",
    "annuity_model/src/annuity_model/crediting.py",
    "annuity_model/src/annuity_model/lapse.py",
    "annuity_model/src/annuity_model/data_registry.py",
    "annuity_model/scripts/agent_preflight.py",
    "annuity_model/scripts/agent_team_router.py",
    "annuity_model/scripts/check_team_run_packet_evidence.py",
    "annuity_model/src/annuity_model/products/**/*.py",
    "annuity_model/src/annuity_model/ui/**/*.py",
)

TEST_GLOBS = (
    "annuity_model/tests/**/*.py",
    "annuity_model/tests/*.py",
)


@dataclass(frozen=True)
class GuardResult:
    behavior_files: tuple[str, ...]
    test_files: tuple[str, ...]

    @property
    def ok(self) -> bool:
        return not self.behavior_files or bool(self.test_files)


def _git_changed_files(base_sha: str | None, head_sha: str | None) -> list[str]:
    if bool(base_sha) != bool(head_sha):
        raise ValueError("Provide both base_sha and head_sha, or neither.")
    if base_sha and head_sha:
        cmd = ["git", "diff", "--name-only", f"{base_sha}..{head_sha}"]
    else:
        cmd = ["git", "diff", "--cached", "--name-only"]
    # Reviewed: fixed git command vector; refs are CI/developer supplied and shell=False.
    out = subprocess.check_output(cmd, cwd=REPO_ROOT, text=True)  # nosec B603
    return [line.strip() for line in out.splitlines() if line.strip()]


def _match_any(path: str, globs: tuple[str, ...]) -> bool:
    return any(fnmatch.fnmatch(path, pattern) for pattern in globs)


def evaluate_changed_files(changed_files: list[str]) -> GuardResult:
    behavior = tuple(sorted(p for p in changed_files if _match_any(p, BEHAVIOR_GLOBS)))
    tests = tuple(sorted(p for p in changed_files if _match_any(p, TEST_GLOBS)))
    return GuardResult(behavior_files=behavior, test_files=tests)


def _print_failure(result: GuardResult) -> None:
    print("Behavior-impacting code changed without matching test updates.")
    print("Changed behavior files:")
    for path in result.behavior_files:
        print(f"  - {path}")
    print(
        "Add or update pytest files under annuity_model/tests/ (or tests/parity/) "
        "in the same change."
    )


def main(argv: list[str]) -> int:
    base_sha = argv[1] if len(argv) >= 2 else None
    head_sha = argv[2] if len(argv) >= 3 else None
    changed = _git_changed_files(base_sha, head_sha)
    result = evaluate_changed_files(changed)
    if result.ok:
        print("OK: unit-test discipline guard passed.")
        return 0
    _print_failure(result)
    return 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
