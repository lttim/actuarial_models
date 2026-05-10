"""RILA contract + result dataclasses (re-export shim).

Canonical import path; implementation lives in :mod:`rila_projection`.
"""

from __future__ import annotations

from annuity_model.rila_projection import (
    RILAContract,
    RILAMonteCarloResult,
    RILAPricingInfeasibleError,
    RILAProjectionResult,
)

__all__ = [
    "RILAContract",
    "RILAMonteCarloResult",
    "RILAPricingInfeasibleError",
    "RILAProjectionResult",
]
