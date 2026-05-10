"""VARIABLE_UL pricing + ALM engine surface (re-export shim)."""

from __future__ import annotations

from annuity_model.vul_projection import (
    VULContract,
    VULProjectionResult,
    liability_path_from_vul_projection,
    price_vul_single_premium,
    price_vul_single_premium_monte_carlo,
)

__all__ = [
    "VULContract",
    "VULProjectionResult",
    "liability_path_from_vul_projection",
    "price_vul_single_premium",
    "price_vul_single_premium_monte_carlo",
]
