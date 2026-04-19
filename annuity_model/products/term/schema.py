"""Term Life contract + result dataclasses (re-export shim).

Canonical import path; implementation lives in :mod:`term_projection`.
"""

from __future__ import annotations

from term_projection import TermLifeContract, TermLifeProjectionResult

__all__ = ["TermLifeContract", "TermLifeProjectionResult"]
