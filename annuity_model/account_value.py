"""Account-value (AV) engine shared by UL / IUL / VUL.

The monthly AV cycle is the same for all three universal-life variants;
only the source of the per-month credit rate differs:

* UL  -> a flat declared rate (constant ``credit_monthly`` across months)
* IUL -> zero except on segment anniversaries, where the credit comes
         from a :class:`crediting.AnnualPointToPointCapped` strategy.
* VUL -> a sub-account return path (deterministic CSV or GBM).

This module captures the cycle once and lets the per-product engines
supply the ``monthly_credit_rate`` and ``monthly_coi_q`` arrays.

Equation
--------

For each month ``t = 0..n_months-1``::

    av_after_premium_load = AV[t] + premium_credit_at_month_t
    av_after_credit       = av_after_premium_load * (1 + credit_monthly[t])
    db_t                  = max(face_amount, av_after_credit)   # Type A
    nar_t                 = max(0, db_t - av_after_credit)
    coi_t                 = monthly_coi_q[t] * nar_t
    av_after_coi          = av_after_credit - coi_t
    av_after_charges      = av_after_coi - monthly_expense_charge
    AV[t+1]               = max(0, av_after_charges)

Premium load is one-time at issue (month 0). Once ``AV[t] == 0`` the
contract terminates; downstream engines are responsible for zeroing
post-termination cashflows. The function returns the per-month AV path
plus the per-month COI dollars and the per-month death benefit dollars
so engines can build cashflow vectors directly.

References
----------
* Section 1.3 of ``docs/seven_product_rollout_plan.md``.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

import numpy as np


@dataclass(frozen=True, slots=True)
class AVConfig:
    """Inputs that don't vary month-by-month."""

    initial_premium: float
    premium_load_pct: float
    monthly_expense_charge: float
    db_type: Literal["return_of_av", "level_face"]
    face_amount: float

    def __post_init__(self) -> None:
        if float(self.initial_premium) < 0.0:
            raise ValueError(
                f"initial_premium must be >= 0; got {self.initial_premium!r}"
            )
        if not (0.0 <= float(self.premium_load_pct) < 1.0):
            raise ValueError(
                f"premium_load_pct must be in [0, 1); got {self.premium_load_pct!r}"
            )
        if float(self.monthly_expense_charge) < 0.0:
            raise ValueError(
                f"monthly_expense_charge must be >= 0; got "
                f"{self.monthly_expense_charge!r}"
            )
        if self.db_type not in ("return_of_av", "level_face"):
            raise ValueError(
                f"db_type must be 'return_of_av' or 'level_face'; got "
                f"{self.db_type!r}"
            )
        if float(self.face_amount) <= 0.0:
            raise ValueError(
                f"face_amount must be > 0; got {self.face_amount!r}"
            )


@dataclass(frozen=True, slots=True)
class AVEvolution:
    """Per-month outputs from :func:`evolve_account_value`.

    All arrays have length ``n_months``.

    Attributes
    ----------
    account_value_end_month:
        ``AV[t+1]`` for each month (end-of-month value after premium
        load, credit, COI, and expense charge).
    db_end_month:
        Death benefit applicable at the end of each month (Type A:
        ``max(face, AV)``; Type B: ``face + AV``).
    coi_dollars:
        Cost-of-insurance dollars deducted that month.
    nar_end_month:
        Net Amount at Risk at end of month (``DB - AV_after_credit``).
    credit_applied:
        The actual credited rate applied that month (echoed back so
        engines can fold it into reporting).
    is_terminated_after_month:
        Boolean array; ``True`` means AV reached 0 at the end of that
        month. Subsequent months stay terminated.
    """

    account_value_end_month: np.ndarray
    db_end_month: np.ndarray
    coi_dollars: np.ndarray
    nar_end_month: np.ndarray
    credit_applied: np.ndarray
    is_terminated_after_month: np.ndarray


def _db_at(av_after_credit: float, face_amount: float, db_type: str) -> float:
    if db_type == "level_face":
        return float(max(face_amount, av_after_credit))
    return float(face_amount + max(0.0, av_after_credit))


def evolve_account_value(
    *,
    config: AVConfig,
    n_months: int,
    monthly_credit_rate: np.ndarray,
    monthly_coi_q: np.ndarray,
) -> AVEvolution:
    """Walk the monthly AV cycle and return per-month vectors.

    Parameters
    ----------
    config:
        Static contract parameters.
    n_months:
        Number of months in the projection.
    monthly_credit_rate:
        Per-month credited rate (decimals, e.g. 0.003 for 30bps). Must
        have length ``n_months``.
    monthly_coi_q:
        Per-month mortality probability used in the COI calculation.
        Must have length ``n_months``.

    Returns
    -------
    :class:`AVEvolution` containing per-month state vectors.

    Notes
    -----
    * Premium load is applied once at month 0 (the entry to the cycle).
    * COI is computed AFTER credit and BEFORE the expense charge.
    * If ``AV`` would go below 0 due to COI + expense, it is floored
      at 0 and the contract is marked terminated for that month and
      every subsequent month.
    """
    if n_months < 0:
        raise ValueError(f"n_months must be >= 0; got {n_months!r}")
    cred = np.asarray(monthly_credit_rate, dtype=float)
    qm = np.asarray(monthly_coi_q, dtype=float)
    if cred.shape != (n_months,):
        raise ValueError(
            f"monthly_credit_rate shape {cred.shape!r} != ({n_months},)"
        )
    if qm.shape != (n_months,):
        raise ValueError(
            f"monthly_coi_q shape {qm.shape!r} != ({n_months},)"
        )

    av_end = np.zeros(n_months, dtype=float)
    db_end = np.zeros(n_months, dtype=float)
    coi_d = np.zeros(n_months, dtype=float)
    nar = np.zeros(n_months, dtype=float)
    credit_applied = np.zeros(n_months, dtype=float)
    terminated = np.zeros(n_months, dtype=bool)

    # AV[0] before any month-1 activity is 0 (premium is loaded inside
    # month 1 below). This matches the industry convention "premium is
    # received at issue, immediately credited, immediately charged".
    av = 0.0
    is_term = False
    initial_load = float(config.initial_premium) * (1.0 - float(config.premium_load_pct))
    for t in range(n_months):
        if is_term:
            av_end[t] = 0.0
            db_end[t] = 0.0
            coi_d[t] = 0.0
            nar[t] = 0.0
            credit_applied[t] = 0.0
            terminated[t] = True
            continue
        # Premium load applied once at month 0 (inside month 1's cycle).
        if t == 0:
            av_after_premium = av + initial_load
        else:
            av_after_premium = av
        cr = float(cred[t])
        av_after_credit = av_after_premium * (1.0 + cr)
        db_t = _db_at(av_after_credit, float(config.face_amount), config.db_type)
        nar_t = max(0.0, db_t - av_after_credit)
        coi_t = float(qm[t]) * nar_t
        av_after_coi = av_after_credit - coi_t
        av_after_charges = av_after_coi - float(config.monthly_expense_charge)
        if av_after_charges <= 0.0:
            av = 0.0
            is_term = True
        else:
            av = av_after_charges
        av_end[t] = float(av)
        db_end[t] = float(db_t)
        coi_d[t] = float(coi_t)
        nar[t] = float(nar_t)
        credit_applied[t] = float(cr)
        terminated[t] = is_term

    return AVEvolution(
        account_value_end_month=av_end,
        db_end_month=db_end,
        coi_dollars=coi_d,
        nar_end_month=nar,
        credit_applied=credit_applied,
        is_terminated_after_month=terminated,
    )


__all__ = [
    "AVConfig",
    "AVEvolution",
    "evolve_account_value",
]
