"""VARIABLE_ANNUITY pricing + ALM engine surface (re-export shim)."""

from __future__ import annotations

from annuity_model.va_projection import (
    VAContract,
    VAProjectionResult,
    liability_path_from_va_projection,
    price_va_single_premium,
    price_va_single_premium_monte_carlo,
)

__all__ = [
    "VAContract",
    "VAProjectionResult",
    "liability_path_from_va_projection",
    "price_va_single_premium",
    "price_va_single_premium_monte_carlo",
]
