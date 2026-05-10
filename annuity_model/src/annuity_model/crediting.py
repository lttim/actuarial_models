"""Crediting-strategy framework for indexed and fixed-rate products.

A small strategy hierarchy that captures the per-segment crediting
arithmetic used by RILA, FIA, IUL (and could be extended to other
products that pay interest tied to an index or a declared rate).

* :class:`CreditingStrategy` is the structural ``Protocol``.
* :class:`FixedDeclaredRate` returns a constant per-segment rate
  regardless of the index path (used by MYGA-style declared-rate
  products and as a UL declared-rate baseline).
* :class:`AnnualPointToPointCapped` implements the participation +
  cap + floor formula used by RILA / FIA / IUL.

The existing RILA inline ``segment_credited_return`` is kept as a thin
wrapper around :class:`AnnualPointToPointCapped` so the public name and
its golden JSON stay byte-identical (Section 1.2 of
``docs/seven_product_rollout_plan.md``).

This module is intentionally standalone (no project imports) so it can
be referenced from any engine without forming an import cycle.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Protocol

from annuity_model.policy_features import buffer_credited_return


class CreditingStrategy(Protocol):
    """Per-segment crediting strategy.

    Implementations consume the *raw* index return for the segment
    (a decimal, e.g. 0.12 for +12%) and return the credited decimal
    that should be applied to the account value for that segment.
    """

    def credit_segment(self, *, raw_index_return: float) -> float:  # pragma: no cover - protocol
        ...


@dataclass(frozen=True, slots=True)
class FixedDeclaredRate:
    """Return a constant per-segment annual rate regardless of the index.

    Used by MYGA, UL declared-rate buckets, and as a unit-test baseline
    for FIA/IUL crediting (cap = floor = annual_rate => collapses to
    fixed).
    """

    annual_rate: float

    def credit_segment(self, *, raw_index_return: float) -> float:
        del raw_index_return
        return float(self.annual_rate)


@dataclass(frozen=True, slots=True)
class AnnualPointToPointCapped:
    """Cap + floor + participation rate, applied to the segment's
    point-to-point index return.

    Returns ``max(floor, min(cap, participation * raw))``. This is the
    strategy used by RILA (with floor potentially negative), FIA
    (typically floor = 0), and IUL (typically floor = 0).

    Constraints
    -----------
    * ``participation >= 0``
    * ``cap >= floor``

    Both are enforced at construction; per-product upstream validators
    additionally ensure participation/cap/floor land in product-specific
    ranges (see ``ProductDefinition.validator``).
    """

    participation: float
    cap: float
    floor: float

    def __post_init__(self) -> None:
        if float(self.participation) < 0.0:
            raise ValueError(f"participation must be >= 0; got {self.participation!r}")
        if float(self.cap) < float(self.floor):
            raise ValueError(f"cap ({self.cap!r}) must be >= floor ({self.floor!r})")

    def credit_segment(self, *, raw_index_return: float) -> float:
        x = float(self.participation) * float(raw_index_return)
        return float(max(float(self.floor), min(float(self.cap), x)))


@dataclass(frozen=True, slots=True)
class AnnualPointToPointBuffer:
    """Annual point-to-point RILA buffer design.

    Positive returns receive participation up to ``cap``. Negative returns
    are protected by ``buffer`` first, so a -15% raw return with a 10% buffer
    credits -5%.
    """

    participation: float
    cap: float
    buffer: float

    def __post_init__(self) -> None:
        if float(self.participation) < 0.0:
            raise ValueError("participation must be >= 0.")
        if float(self.cap) < 0.0:
            raise ValueError("cap must be >= 0 for buffer designs.")
        if not (0.0 <= float(self.buffer) <= 1.0):
            raise ValueError("buffer must be in [0, 1].")

    def credit_segment(self, *, raw_index_return: float) -> float:
        return buffer_credited_return(
            raw_index_return=float(raw_index_return),
            participation=float(self.participation),
            cap=float(self.cap),
            buffer=float(self.buffer),
        )


def segment_credited_return_from_strategy(
    *, strategy: CreditingStrategy, raw_index_return: float
) -> float:
    """Convenience wrapper for callers that prefer a free function.

    Mirrors :meth:`CreditingStrategy.credit_segment` so a function-only
    call site (e.g. an Excel-formula generator translating Python logic
    one-to-one) can use either form.
    """
    return float(strategy.credit_segment(raw_index_return=float(raw_index_return)))


__all__ = [
    "AnnualPointToPointBuffer",
    "AnnualPointToPointCapped",
    "CreditingStrategy",
    "FixedDeclaredRate",
    "segment_credited_return_from_strategy",
]
