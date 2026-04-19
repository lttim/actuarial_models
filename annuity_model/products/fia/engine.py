    """FIA pricing + ALM engine surface (re-export shim)."""

    from __future__ import annotations

    from fia_projection import (
        FIAContract,
FIAProjectionResult,
liability_path_from_fia_projection,
price_fia_single_premium,
price_fia_single_premium_monte_carlo,
    )

    __all__ = [
    "FIAContract",
    "FIAProjectionResult",
    "liability_path_from_fia_projection",
    "price_fia_single_premium",
    "price_fia_single_premium_monte_carlo",
]
