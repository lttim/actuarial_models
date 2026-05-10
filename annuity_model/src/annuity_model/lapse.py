"""Static lapse / persistency framework.

A small, opt-in module every new product engine can plug into. The
contract is intentionally minimal:

* :class:`LapseAssumption` carries a tuple of annual policy-year lapse
  rates plus an *ultimate* rate for years beyond the table. ``None``
  remains the engine-side "no lapse" sentinel everywhere.
* :func:`combined_monthly_survival` composes mortality and lapse
  multiplicatively under independence.
* :func:`monthly_decrements` converts an annual ``q_w`` to a flat
  monthly hazard via ``1 - (1 - q_w)**(1/12)`` (constant force inside
  the policy year, same convention used for mortality elsewhere).

This module is intentionally standalone (no project imports) so it can
be referenced from any engine without forming an import cycle. Existing
SPIA / Term / RILA engines stay mortality-only; the optional ``lapse=``
slot is added for future use but defaults to ``None`` everywhere
(verified by the unchanged SPIA / Term / RILA golden JSON).

References
----------
* Section 1.1 of ``docs/seven_product_rollout_plan.md``.
* ``docs/lapse_framework.md`` for the framework rationale.
"""

from __future__ import annotations

from dataclasses import dataclass

import numpy as np


@dataclass(frozen=True, slots=True)
class LapseAssumption:
    """Per-policy-year annual lapse table plus an ultimate fallback.

    Attributes
    ----------
    annual_lapse_rates_by_year:
        ``q_w[y]`` for policy year ``y`` (1-indexed but stored
        0-indexed; entry 0 is year 1, entry 1 is year 2, ...).
        Each element must satisfy ``0 <= q_w < 1``.
    ultimate_rate:
        The annual lapse rate used after the table runs out (i.e. for
        policy year ``len(annual_lapse_rates_by_year) + 1`` and beyond).
        Must satisfy ``0 <= ultimate_rate < 1``.
    """

    annual_lapse_rates_by_year: tuple[float, ...]
    ultimate_rate: float = 0.0

    def __post_init__(self) -> None:
        for i, q in enumerate(self.annual_lapse_rates_by_year):
            if not (0.0 <= float(q) < 1.0):
                raise ValueError(
                    f"annual_lapse_rates_by_year[{i}]={q!r} must be in [0, 1); "
                    "lapse rates of 1.0 mean immediate certain lapse and are "
                    "almost always a unit error."
                )
        if not (0.0 <= float(self.ultimate_rate) < 1.0):
            raise ValueError(f"ultimate_rate={self.ultimate_rate!r} must be in [0, 1).")

    def annual_rate_for_policy_year(self, policy_year: int) -> float:
        """Return ``q_w`` for *policy_year* (1-indexed).

        Falls back to :attr:`ultimate_rate` after the table runs out.
        """
        if policy_year <= 0:
            raise ValueError(f"policy_year must be >= 1; got {policy_year!r}")
        idx = policy_year - 1
        if idx < len(self.annual_lapse_rates_by_year):
            return float(self.annual_lapse_rates_by_year[idx])
        return float(self.ultimate_rate)

    def monthly_decrements(self, n_months: int) -> np.ndarray:
        """Return ``q_w_m`` for each of *n_months* months (length ``n_months``).

        Inside policy year ``y`` we apply a flat monthly hazard
        ``1 - (1 - q_w[y])**(1/12)`` -- constant force within the
        policy year, same convention as monthly mortality.
        """
        if n_months < 0:
            raise ValueError(f"n_months must be >= 0; got {n_months!r}")
        out = np.zeros(n_months, dtype=float)
        for m in range(n_months):
            policy_year = (m // 12) + 1
            q_annual = self.annual_rate_for_policy_year(policy_year)
            q_monthly = 1.0 - (1.0 - q_annual) ** (1.0 / 12.0)
            out[m] = q_monthly
        return out


def combined_monthly_survival(
    *,
    mortality_monthly_q: np.ndarray,
    lapse_monthly_q: np.ndarray,
) -> np.ndarray:
    """Return ``S(t)`` under independent mortality + lapse decrements.

    ``S(t) = ∏_{s=0}^{t-1} (1 - q_x_m(s)) * (1 - q_w_m(s))``.

    Both inputs must have the same length. The output has the same
    length and represents survival to the *end* of each month
    (``S[0]`` is survival through month 1, ``S[1]`` is survival
    through months 1 and 2, etc.).
    """
    qm = np.asarray(mortality_monthly_q, dtype=float)
    qw = np.asarray(lapse_monthly_q, dtype=float)
    if qm.shape != qw.shape:
        raise ValueError(
            f"mortality_monthly_q shape {qm.shape!r} must equal lapse_monthly_q shape {qw.shape!r}."
        )
    if qm.ndim != 1:
        raise ValueError("inputs must be 1-D arrays.")
    qm = np.clip(qm, 0.0, 1.0)
    qw = np.clip(qw, 0.0, 1.0)
    one_minus = (1.0 - qm) * (1.0 - qw)
    return np.cumprod(one_minus)


def monthly_mortality_q_from_annual(annual_qx: np.ndarray) -> np.ndarray:
    """Convert an array of annual ``q_x`` values to monthly ``q_x_m``.

    Uses the constant-force-of-mortality convention
    ``q_m = 1 - (1 - q_annual)**(1/12)``. Mirrors the same conversion
    used by :class:`pricing_projection.MortalityTableQx`.
    """
    qa = np.clip(np.asarray(annual_qx, dtype=float), 0.0, 0.999999)
    return 1.0 - (1.0 - qa) ** (1.0 / 12.0)


def default_lapse_assumption() -> LapseAssumption:
    """Industry-pattern declining-then-ultimate lapse table.

    Returns
    -------
    A :class:`LapseAssumption` with annual rates 8/7/6/5/4/3/2 percent
    in years 1-7 and an ultimate rate of 2 percent thereafter. This is
    a generic placeholder; production users should override with their
    own table.
    """
    return LapseAssumption(
        annual_lapse_rates_by_year=(0.08, 0.07, 0.06, 0.05, 0.04, 0.03, 0.02),
        ultimate_rate=0.02,
    )


def lapse_decrement_from_csv(path: str) -> LapseAssumption:
    """Load a lapse table from a CSV with columns ``policy_year, q_w``.

    Years must be 1-indexed and contiguous. Any trailing row with
    ``policy_year`` blank is treated as the ultimate-rate row.
    """
    import csv

    rates: list[tuple[int, float]] = []
    ultimate: float = 0.0
    with open(path, encoding="utf-8") as fh:
        reader = csv.DictReader(fh)
        for row in reader:
            year_str = (row.get("policy_year") or "").strip()
            q_str = (row.get("q_w") or "").strip()
            if not year_str:
                if q_str:
                    ultimate = float(q_str)
                continue
            rates.append((int(year_str), float(q_str)))
    rates.sort(key=lambda x: x[0])
    if rates:
        years, vals = zip(*rates, strict=False)
        if list(years) != list(range(1, len(rates) + 1)):
            raise ValueError(
                f"lapse CSV {path!r} must have contiguous 1-indexed policy_year "
                f"rows; got {list(years)!r}."
            )
        return LapseAssumption(
            annual_lapse_rates_by_year=tuple(float(v) for v in vals),
            ultimate_rate=float(ultimate),
        )
    return LapseAssumption(annual_lapse_rates_by_year=(), ultimate_rate=float(ultimate))


__all__ = [
    "LapseAssumption",
    "combined_monthly_survival",
    "default_lapse_assumption",
    "lapse_decrement_from_csv",
    "monthly_mortality_q_from_annual",
]
