"""Term Life pricing + ALM engine surface (re-export shim)."""

from __future__ import annotations

from annuity_model.term_projection import (
    TermLifeContract,
    TermLifeProjectionResult,
    liability_path_from_term_projection,
    price_term_life_level_monthly,
)

__all__ = [
    "TermLifeContract",
    "TermLifeProjectionResult",
    "liability_path_from_term_projection",
    "price_term_life_level_monthly",
]
