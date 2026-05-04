#!/usr/bin/env python3
"""Fail when ``parity_constants.py`` changes without a matching
``model_change_log.md`` entry in the same PR.

Background
----------
Parity tolerances are the platform's load-bearing numerical contract with
Excel and with downstream actuarial users. Loosening one is a model decision
that requires a written justification, attached parity-trace before/after,
and a reviewer sign-off. ``CONTRIBUTING.md``, ``parity_test_checklist.md``,
and ``release.md`` all document this; this script is the *mechanical*
enforcement so the rule cannot be skipped.

Usage
-----
::

    python scripts/check_parity_constants_changelog.py BASE_REF [HEAD_REF]

``BASE_REF`` defaults to ``origin/main`` if not provided. ``HEAD_REF``
defaults to ``HEAD``. CI typically calls it with the GitHub-provided
``${{ github.event.pull_request.base.sha }}`` so the comparison is always
against the PR's true base, not whatever ``main`` happens to be at run time.

Exit codes
----------
* ``0`` -- either ``parity_constants.py`` did not change, OR it changed AND
  ``model_change_log.md`` also changed in the same PR.
* ``1`` -- ``parity_constants.py`` changed without a matching log entry.
* ``2`` -- usage error / could not invoke git.

This script imports nothing from the project (it runs standalone in CI
before deps are installed) and uses only the standard library.
"""

from __future__ import annotations

# Reviewed: CI guard invokes a fixed local git command with shell=False.
import subprocess  # nosec B404
import sys
from pathlib import Path

_PARITY_CONSTANTS = "annuity_model/parity_constants.py"
_MODEL_CHANGE_LOG = "annuity_model/docs/model_change_log.md"


def _changed_files(base_ref: str, head_ref: str) -> set[str]:
    """Return the set of paths changed between *base_ref* and *head_ref*.

    ``git diff --name-only base...head`` shows files in the symmetric
    difference of the two refs, which is what GitHub also uses for PR
    file lists. We deliberately do NOT use ``base..head`` (two dots) because
    that misses files added on main after the PR branched.
    """
    try:
        # Reviewed: fixed git command vector; refs are CI/developer supplied and shell=False.
        out = subprocess.check_output(  # nosec B603 B607
            ["git", "diff", "--name-only", f"{base_ref}...{head_ref}"],
            text=True,
        )
    except FileNotFoundError:
        print("error: git is not on PATH", file=sys.stderr)
        sys.exit(2)
    except subprocess.CalledProcessError as exc:
        print(
            f"error: git diff failed (exit {exc.returncode}); "
            f"is {base_ref!r} fetched in this checkout?",
            file=sys.stderr,
        )
        sys.exit(2)
    return {line.strip() for line in out.splitlines() if line.strip()}


def main(argv: list[str]) -> int:
    if len(argv) < 2 or argv[1] in {"-h", "--help"}:
        print(__doc__)
        return 2 if len(argv) < 2 else 0
    base_ref = argv[1]
    head_ref = argv[2] if len(argv) > 2 else "HEAD"

    repo_root = Path(__file__).resolve().parents[2]
    if not (repo_root / ".git").exists():
        print(
            f"error: {repo_root} is not the repo root (no .git dir found)",
            file=sys.stderr,
        )
        return 2

    changed = _changed_files(base_ref, head_ref)
    constants_touched = _PARITY_CONSTANTS in changed
    log_touched = _MODEL_CHANGE_LOG in changed

    if not constants_touched:
        print(f"OK: {_PARITY_CONSTANTS} not modified in this PR; nothing to enforce.")
        return 0

    if log_touched:
        print(f"OK: {_PARITY_CONSTANTS} changed AND {_MODEL_CHANGE_LOG} also changed.")
        return 0

    print(
        "FAIL: "
        f"{_PARITY_CONSTANTS} was modified but {_MODEL_CHANGE_LOG} was not.\n"
        "\n"
        "Tolerance / parity-constant changes require a written justification "
        "with a parity-trace before/after attached. Add an entry to "
        f"{_MODEL_CHANGE_LOG} (see annuity_model/docs/runbooks/release.md "
        "and annuity_model/docs/parity_test_checklist.md), then push again.\n"
        "\n"
        "If you genuinely believe this constants change does NOT affect any "
        "model output (e.g. comment-only edit), add an explicit "
        "'No-op constants change' line to the model change log so the "
        "decision is recorded.",
        file=sys.stderr,
    )
    return 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
