from __future__ import annotations

import datetime as dt
from pathlib import Path

from scripts.check_assumption_release_guardrails import validate_waiver_file

FLAGGED = [
    "expenses_assumptions_us_placeholders (expenses/us_placeholders)",
    "cso_2017_ult_male_nonsmoker_qx (mortality/cso_2017_ult)",
]


def _write_waiver(path: Path, *, expiry: str = "2026-12-31") -> None:
    path.write_text(
        "\n".join(
            [
                "- **Release version:** demo-current",
                "- **Date:** 2026-05-03",
                "- **Approved by:** Demo governance steward",
                "- **Independent challenger:** Actuarial peer review agent",
                (
                    "- **Artifacts covered:** expenses_assumptions_us_placeholders; "
                    "cso_2017_ult_male_nonsmoker_qx"
                ),
                "- **Business justification:** Demo use without licensed production artifacts.",
                "- **Compensating controls:** Visible warnings and release guardrails.",
                f"- **Expiry date for waiver:** {expiry}",
                "",
            ]
        ),
        encoding="utf-8",
    )


def test_valid_assumption_waiver_passes_metadata_checks(tmp_path: Path) -> None:
    waiver = tmp_path / "assumption_waiver.md"
    _write_waiver(waiver)

    assert (
        validate_waiver_file(
            waiver,
            flagged_artifacts=FLAGGED,
            today=dt.date(2026, 5, 3),
        )
        == []
    )


def test_assumption_waiver_rejects_blank_required_fields(tmp_path: Path) -> None:
    waiver = tmp_path / "assumption_waiver.md"
    _write_waiver(waiver)
    text = waiver.read_text(encoding="utf-8")
    waiver.write_text(
        text.replace("- **Approved by:** Demo governance steward", "- **Approved by:**"),
        encoding="utf-8",
    )

    errors = validate_waiver_file(
        waiver,
        flagged_artifacts=FLAGGED,
        today=dt.date(2026, 5, 3),
    )

    assert "waiver field is blank or missing: Approved by" in errors


def test_assumption_waiver_rejects_expired_approval(tmp_path: Path) -> None:
    waiver = tmp_path / "assumption_waiver.md"
    _write_waiver(waiver, expiry="2026-01-01")

    errors = validate_waiver_file(
        waiver,
        flagged_artifacts=FLAGGED,
        today=dt.date(2026, 5, 3),
    )

    assert any("waiver expired" in error for error in errors)
