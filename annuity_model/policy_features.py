"""Shared policy-mechanics primitives for richer RILA / IUL projections.

The classes here are deliberately product-neutral and month-indexed. Engines
own product-specific timing, but schedules and rider/loan arithmetic live in
one place so Python, Excel builders, and tests can use the same vocabulary.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

import numpy as np


CreditDesign = Literal["cap_floor", "buffer"]


@dataclass(frozen=True, slots=True)
class MonthlySchedule:
    """A deterministic dollar schedule indexed by projection month.

    ``amounts[k]`` applies in policy month ``k + 1``. Missing tail months
    are padded with zero.
    """

    amounts: tuple[float, ...] = ()

    def values(self, n_months: int) -> np.ndarray:
        if n_months < 0:
            raise ValueError("n_months must be >= 0.")
        out = np.zeros(int(n_months), dtype=float)
        if not self.amounts:
            return out
        raw = np.asarray(self.amounts, dtype=float)
        if raw.ndim != 1:
            raise ValueError("MonthlySchedule amounts must be one-dimensional.")
        if np.any(~np.isfinite(raw)) or np.any(raw < 0.0):
            raise ValueError("MonthlySchedule amounts must be finite and non-negative.")
        out[: min(out.size, raw.size)] = raw[: out.size]
        return out


@dataclass(frozen=True, slots=True)
class LevelPremiumSchedule:
    """Level planned premium paid at a regular modal frequency."""

    modal_premium: float = 0.0
    mode_months: int = 12
    start_month: int = 1
    end_month: int | None = None

    def values(self, n_months: int) -> np.ndarray:
        if n_months < 0:
            raise ValueError("n_months must be >= 0.")
        if float(self.modal_premium) < 0.0 or not np.isfinite(float(self.modal_premium)):
            raise ValueError("modal_premium must be finite and non-negative.")
        if int(self.mode_months) < 1:
            raise ValueError("mode_months must be >= 1.")
        if int(self.start_month) < 1:
            raise ValueError("start_month must be >= 1.")
        end = int(self.end_month) if self.end_month is not None else int(n_months)
        if end < int(self.start_month):
            raise ValueError("end_month must be >= start_month when provided.")
        out = np.zeros(int(n_months), dtype=float)
        for month in range(int(self.start_month), min(end, int(n_months)) + 1):
            if (month - int(self.start_month)) % int(self.mode_months) == 0:
                out[month - 1] = float(self.modal_premium)
        return out


@dataclass(frozen=True, slots=True)
class SurrenderChargeSchedule:
    """Annual surrender charge rates by policy year."""

    annual_rates: tuple[float, ...] = ()

    def monthly_rates(self, n_months: int) -> np.ndarray:
        if n_months < 0:
            raise ValueError("n_months must be >= 0.")
        raw = np.asarray(self.annual_rates, dtype=float)
        if raw.size and (np.any(~np.isfinite(raw)) or np.any(raw < 0.0) or np.any(raw > 1.0)):
            raise ValueError("surrender charge rates must be finite and in [0, 1].")
        out = np.zeros(int(n_months), dtype=float)
        for i in range(int(n_months)):
            y = i // 12
            out[i] = float(raw[y]) if y < raw.size else 0.0
        return out


@dataclass(frozen=True, slots=True)
class SegmentAllocation:
    """One index-crediting allocation for a RILA/IUL segment."""

    weight: float = 1.0
    design: CreditDesign = "cap_floor"
    participation: float = 1.0
    cap: float = 0.10
    floor: float = 0.0
    buffer: float = 0.10

    def __post_init__(self) -> None:
        if not np.isfinite(float(self.weight)) or float(self.weight) < 0.0:
            raise ValueError("segment allocation weight must be finite and non-negative.")
        if self.design not in ("cap_floor", "buffer"):
            raise ValueError("segment design must be 'cap_floor' or 'buffer'.")
        if float(self.participation) < 0.0:
            raise ValueError("participation must be non-negative.")
        if float(self.cap) < float(self.floor):
            raise ValueError("cap must be >= floor.")
        if not (0.0 <= float(self.buffer) <= 1.0):
            raise ValueError("buffer must be in [0, 1].")


def normalize_segment_allocations(
    allocations: tuple[SegmentAllocation, ...],
) -> tuple[SegmentAllocation, ...]:
    if not allocations:
        raise ValueError("At least one segment allocation is required.")
    total = float(sum(float(a.weight) for a in allocations))
    if total <= 0.0:
        raise ValueError("Segment allocation weights must sum to a positive value.")
    return tuple(
        SegmentAllocation(
            weight=float(a.weight) / total,
            design=a.design,
            participation=float(a.participation),
            cap=float(a.cap),
            floor=float(a.floor),
            buffer=float(a.buffer),
        )
        for a in allocations
    )


@dataclass(frozen=True, slots=True)
class GLWBRider:
    """Simple GLWB rider: roll-up, annual ratchet, then level withdrawals."""

    enabled: bool = False
    fee_annual: float = 0.0
    rollup_annual: float = 0.0
    withdrawal_rate: float = 0.0
    income_start_month: int = 10**9
    ratchet: bool = True

    def __post_init__(self) -> None:
        for name, value in (
            ("fee_annual", self.fee_annual),
            ("rollup_annual", self.rollup_annual),
            ("withdrawal_rate", self.withdrawal_rate),
        ):
            v = float(value)
            if not np.isfinite(v) or v < 0.0 or v > 1.0:
                raise ValueError(f"{name} must be finite and in [0, 1].")
        if int(self.income_start_month) < 1:
            raise ValueError("income_start_month must be >= 1.")


@dataclass(frozen=True, slots=True)
class LoanTerms:
    """Fixed-rate policy loan terms for IUL-style projections."""

    annual_rate: float = 0.0
    draws: MonthlySchedule = MonthlySchedule()
    repayments: MonthlySchedule = MonthlySchedule()

    def monthly_rate(self) -> float:
        r = float(self.annual_rate)
        if not np.isfinite(r) or r < 0.0:
            raise ValueError("loan annual_rate must be finite and non-negative.")
        return float((1.0 + r) ** (1.0 / 12.0) - 1.0)


def buffer_credited_return(
    *,
    raw_index_return: float,
    participation: float,
    cap: float,
    buffer: float,
) -> float:
    """RILA buffer design: upside participates to cap; downside absorbs buffer first."""

    raw = float(raw_index_return)
    part = float(participation)
    cp = float(cap)
    bf = float(buffer)
    if part < 0.0:
        raise ValueError("participation must be non-negative.")
    if cp < 0.0:
        raise ValueError("cap must be non-negative for buffer designs.")
    if not (0.0 <= bf <= 1.0):
        raise ValueError("buffer must be in [0, 1].")
    if raw >= 0.0:
        return float(min(cp, part * raw))
    return float(min(0.0, raw + bf))


def segment_credited_return(*, allocation: SegmentAllocation, raw_index_return: float) -> float:
    if allocation.design == "buffer":
        return buffer_credited_return(
            raw_index_return=float(raw_index_return),
            participation=float(allocation.participation),
            cap=float(allocation.cap),
            buffer=float(allocation.buffer),
        )
    x = float(allocation.participation) * float(raw_index_return)
    return float(max(float(allocation.floor), min(float(allocation.cap), x)))


__all__ = [
    "CreditDesign",
    "GLWBRider",
    "LevelPremiumSchedule",
    "LoanTerms",
    "MonthlySchedule",
    "SegmentAllocation",
    "SurrenderChargeSchedule",
    "buffer_credited_return",
    "normalize_segment_allocations",
    "segment_credited_return",
]
