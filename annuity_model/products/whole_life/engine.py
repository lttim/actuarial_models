"""WHOLE_LIFE pricing + ALM engine surface (re-export shim)."""

from __future__ import annotations

from wl_projection import (
    WLContract,
    WLProjectionResult,
    liability_path_from_wl_projection,
    price_wl_single_premium,
)

__all__ = [
    "WLContract",
    "WLProjectionResult",
    "liability_path_from_wl_projection",
    "price_wl_single_premium",
]
