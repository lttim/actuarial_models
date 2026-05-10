"""SPIA pricing + ALM engine surface (re-export shim).

The deterministic and Monte Carlo pricing entry points, the
``LiabilityPath`` converter, and the contract/result dataclasses are all
exposed here under the canonical ``products.spia.engine`` path. The
implementation continues to live in :mod:`pricing_projection` until a
later wave physically moves it; the parity of names is enforced by
:mod:`tests.test_products_subpackage_shims`.
"""

from __future__ import annotations

from annuity_model.pricing_projection import (
    SPIAContract,
    SPIAMonteCarloResult,
    SPIAProjectionResult,
    liability_path_from_spia_projection,
    price_spia_single_premium,
    price_spia_single_premium_monte_carlo,
)

__all__ = [
    "SPIAContract",
    "SPIAMonteCarloResult",
    "SPIAProjectionResult",
    "liability_path_from_spia_projection",
    "price_spia_single_premium",
    "price_spia_single_premium_monte_carlo",
]
