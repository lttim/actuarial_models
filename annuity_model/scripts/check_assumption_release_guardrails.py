"""Release guardrail checks for placeholder/synthetic assumptions.

This script fails release readiness when assumption artifacts flagged as
synthetic/placeholder are present without an explicit waiver file.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

_HERE = Path(__file__).resolve()
_REPO_ROOT = _HERE.parent.parent
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

from data_registry import REGISTRY


def _looks_non_production_source(source: str) -> bool:
    s = source.lower()
    needles = (
        "synthetic placeholder",
        "placeholder",
        "not licensed",
        "development only",
    )
    return any(n in s for n in needles)


def _collect_flagged_artifacts() -> list[str]:
    flagged: list[str] = []
    for artifact in REGISTRY:
        if _looks_non_production_source(artifact.source):
            flagged.append(f"{artifact.name} ({artifact.kind}/{artifact.version})")
    return flagged


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Fail if placeholder assumptions are present without waiver."
    )
    parser.add_argument(
        "--waiver-file",
        default=".release/assumption_waiver.md",
        help=(
            "Path to waiver file relative to annuity_model/. "
            "Required when placeholder assumptions are detected."
        ),
    )
    args = parser.parse_args()

    flagged = _collect_flagged_artifacts()
    if not flagged:
        print("PASS: no placeholder assumptions detected in data_registry.")
        return 0

    waiver_path = Path(args.waiver_file)
    if waiver_path.exists():
        print("PASS (with waiver): placeholder assumptions detected and waiver file exists:")
        for row in flagged:
            print(f"  - {row}")
        print(f"Waiver: {waiver_path}")
        return 0

    print("FAIL: placeholder assumptions detected with no waiver file.")
    for row in flagged:
        print(f"  - {row}")
    print("")
    print("To proceed, either:")
    print("  1) Replace placeholder assumptions with approved production artifacts, or")
    print(f"  2) Provide waiver evidence at {waiver_path} (see docs/assumption_governance.md).")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
