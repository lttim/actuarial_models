"""SPIA contract + result dataclasses (re-export shim).

Today these classes live in :mod:`pricing_projection`; this shim is the
canonical import path that contract scaffolding tooling and new tests
SHOULD use:

.. code-block:: python

    from annuity_model.products.spia.schema import SPIAContract, SPIAProjectionResult

When the implementation later moves into this subpackage, only the
re-exports here change -- callers stay put.
"""

from __future__ import annotations

from annuity_model.pricing_projection import (
    SPIAContract,
    SPIAMonteCarloResult,
    SPIAProjectionResult,
)

__all__ = [
    "SPIAContract",
    "SPIAMonteCarloResult",
    "SPIAProjectionResult",
]
