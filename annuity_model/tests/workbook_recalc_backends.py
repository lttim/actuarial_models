"""Test-only interfaces for experimental workbook recalc backends.

The mandatory Excel gates in this repo are static/workbook-contract tests.
Runtime recalc candidates live behind this tiny Protocol so a future backend
can be evaluated against the corpus without becoming a merge-blocking
dependency by accident.
"""

from __future__ import annotations

from collections.abc import Sequence
from typing import Protocol, TypeAlias

WorkbookCellValue: TypeAlias = float | int | str | None


class WorkbookRecalcBackend(Protocol):
    """Experimental workbook formula evaluator used by advisory tests only."""

    name: str

    def recalc(
        self,
        raw_xlsx: bytes,
        cells: Sequence[str],
        *,
        timeout: float,
    ) -> dict[str, WorkbookCellValue]:
        """Return recalculated cell values keyed by ``Sheet!A1`` coordinate."""


def candidate_backends() -> tuple[WorkbookRecalcBackend, ...]:
    """Return installed advisory recalc backends.

    Empty by design after the LibreOffice removal. Candidate integrations
    should be added here only after their dependency footprint is acceptable
    for local advisory use; mandatory CI promotion requires separate review.
    """

    return ()
