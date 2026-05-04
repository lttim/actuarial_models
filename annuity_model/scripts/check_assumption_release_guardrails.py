"""Release guardrail checks for placeholder/synthetic assumptions.

This script fails release readiness when assumption artifacts flagged as
synthetic/placeholder are present without an explicit waiver file.
"""

from __future__ import annotations

import argparse
import datetime as dt
import sys
from pathlib import Path

_HERE = Path(__file__).resolve()
_REPO_ROOT = _HERE.parent.parent
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

from data_registry import REGISTRY

REQUIRED_WAIVER_FIELDS = (
    "Release version",
    "Date",
    "Approved by",
    "Independent challenger",
    "Artifacts covered",
    "Business justification",
    "Compensating controls",
    "Expiry date for waiver",
)


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


def _field_value(text: str, field: str) -> str:
    prefix = f"- **{field}:**"
    for line in text.splitlines():
        stripped = line.strip()
        if stripped.startswith(prefix):
            return stripped.removeprefix(prefix).strip()
    return ""


def validate_waiver_file(
    waiver_path: Path,
    *,
    flagged_artifacts: list[str],
    today: dt.date | None = None,
) -> list[str]:
    """Return validation errors for a placeholder-assumption release waiver."""
    if not waiver_path.exists():
        return [f"waiver file does not exist: {waiver_path}"]

    text = waiver_path.read_text(encoding="utf-8")
    errors: list[str] = []
    for field in REQUIRED_WAIVER_FIELDS:
        if not _field_value(text, field):
            errors.append(f"waiver field is blank or missing: {field}")

    for artifact in flagged_artifacts:
        artifact_name = artifact.split(" ", 1)[0]
        if artifact_name not in text:
            errors.append(f"waiver does not list flagged artifact: {artifact_name}")

    expiry_raw = _field_value(text, "Expiry date for waiver")
    if expiry_raw:
        try:
            expiry = dt.date.fromisoformat(expiry_raw)
        except ValueError:
            errors.append("waiver expiry must be ISO date YYYY-MM-DD")
        else:
            effective_today = today or dt.datetime.now(dt.UTC).date()
            if expiry < effective_today:
                errors.append(
                    f"waiver expired on {expiry.isoformat()} (today {effective_today.isoformat()})"
                )
    return errors


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
    parser.add_argument(
        "--today",
        default=None,
        help="Override today's date for tests, as YYYY-MM-DD.",
    )
    args = parser.parse_args()

    flagged = _collect_flagged_artifacts()
    if not flagged:
        print("PASS: no placeholder assumptions detected in data_registry.")
        return 0

    waiver_path = Path(args.waiver_file)
    if not waiver_path.is_absolute():
        waiver_path = _REPO_ROOT / waiver_path
    today = dt.date.fromisoformat(args.today) if args.today else None
    waiver_errors = validate_waiver_file(
        waiver_path,
        flagged_artifacts=flagged,
        today=today,
    )
    if not waiver_errors:
        print("PASS (with waiver): placeholder assumptions detected and waiver file exists:")
        for row in flagged:
            print(f"  - {row}")
        print(f"Waiver: {waiver_path}")
        return 0

    print("FAIL: placeholder assumptions detected without a valid waiver.")
    for row in flagged:
        print(f"  - {row}")
    print("")
    print("Waiver validation errors:")
    for error in waiver_errors:
        print(f"  - {error}")
    print("")
    print("To proceed, either:")
    print("  1) Replace placeholder assumptions with approved production artifacts, or")
    print(f"  2) Provide waiver evidence at {waiver_path} (see docs/assumption_governance.md).")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
