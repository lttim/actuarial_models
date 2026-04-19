    """UNIVERSAL_LIFE pricing + ALM engine surface (re-export shim)."""

    from __future__ import annotations

    from ul_projection import (
        ULContract,
ULProjectionResult,
liability_path_from_ul_projection,
price_ul_single_premium,
    )

    __all__ = [
    "ULContract",
    "ULProjectionResult",
    "liability_path_from_ul_projection",
    "price_ul_single_premium",
]
