"""Per-product actuarial benchmark bands (single source of truth).

The bands here back the "actuarial reasonableness gate" tests
(``tests/parity/test_<P>_actuarial.py``). They are intentionally WIDE
(real-world product pricing depends on assumption choices) but tight
enough to catch order-of-magnitude bugs that pure parity tests cannot
see (Section 13 of ``docs/seven_product_rollout_plan.md``).

Discipline (Section 13.9 of the rollout plan)
---------------------------------------------
* **Tests import constants by name** -- never inline a literal in
  ``tests/parity/test_<P>_actuarial.py``.
* **Failing a band is investigated, never widened** without an
  explanatory paragraph appended to ``docs/actuarial_benchmarks.md``
  ("Band change log" section) and the same review path as
  ``parity_constants.py``.
* **Band rationale lives in** ``docs/actuarial_benchmarks.md``; this
  file is the executable side. The cross-check
  ``scripts/render_actuarial_benchmarks.py --check`` keeps the two in
  sync.

Naming convention
-----------------
* ``<PRODUCT>_BENCHMARK_<METRIC>_LO`` / ``_HI`` for the wide band.
* ``<PRODUCT>_SENSITIVITY_EPS`` for sign-direction tolerance epsilons.
* ``<PRODUCT>_CLOSED_FORM_<KIND>_TOL`` for closed-form match tolerances.
"""

from __future__ import annotations

from typing import Final

# ---------------------------------------------------------------------------
# MYGA: $100k SP, 4.5% rate, 5y guarantee, age 60.
# ---------------------------------------------------------------------------

MYGA_BENCHMARK_AV_T_LO: Final[float] = 124_500.0
"""Lower bound for MYGA account value at maturity.

Closed form: $100,000 * 1.045^5 = $124,618. Band ±~$120 accounts for
discrete monthly compounding."""

MYGA_BENCHMARK_AV_T_HI: Final[float] = 124_800.0
"""Upper bound for MYGA account value at maturity (see *_LO note)."""

MYGA_BENCHMARK_PV_LO: Final[float] = 99_000.0
"""Lower bound for MYGA PV(maturity payout) at flat 4.5% discount.

If discount rate equals declared rate, PV ≈ premium × survival(T)."""

MYGA_BENCHMARK_PV_HI: Final[float] = 101_000.0
"""Upper bound for MYGA PV(maturity payout) at flat 4.5% discount."""

MYGA_CLOSED_FORM_AV_TOL: Final[float] = 1e-2
"""MYGA AV(T) vs SP*(1+i)^T closed form tolerance (USD)."""

MYGA_SENSITIVITY_EPS: Final[float] = 1.0
"""MYGA sensitivity sign tolerance (USD)."""

# ---------------------------------------------------------------------------
# FIA: $100k SP, 80% pop, 7% cap, 0% floor, 10y, age 60, S&P baseline.
# ---------------------------------------------------------------------------

FIA_BENCHMARK_AV_T_LO: Final[float] = 100_000.0
"""Lower bound for FIA AV at horizon (floor 0 -> AV cannot decrease)."""

FIA_BENCHMARK_AV_T_HI: Final[float] = 200_000.0
"""Upper bound for FIA AV at horizon.

Cap 0.07 * 0.8 part = up to 5.6%/yr; 10y compounded ≈ 73% upside."""

FIA_SENSITIVITY_EPS: Final[float] = 1.0
"""FIA sensitivity sign tolerance (USD)."""

# ---------------------------------------------------------------------------
# VA: $100k SP, 6% drift, 1.4% M&E, 20y, age 55.
# ---------------------------------------------------------------------------

VA_BENCHMARK_AV_T_FLAT_LO: Final[float] = 60_000.0
"""Lower bound for VA AV at horizon under flat-S&P deterministic path.

20y of 1.4% M&E charges with 0% return shrinks $100k to ~$75k; band
allows for slight S&P drift in the baseline scenario."""

VA_BENCHMARK_AV_T_FLAT_HI: Final[float] = 110_000.0
"""Upper bound for VA AV at horizon under flat-S&P deterministic path."""

VA_BENCHMARK_AV_T_MC_LO: Final[float] = 170_000.0
"""Lower bound for VA E[AV(T)] under Monte Carlo with 6% drift.

Lognormal moment: exp((0.06-0.014)*20) ≈ 2.51 → ~$251k MC mean.
Lower band allows for stochastic noise in modest n_sims."""

VA_BENCHMARK_AV_T_MC_HI: Final[float] = 320_000.0
"""Upper bound for VA E[AV(T)] under Monte Carlo with 6% drift."""

VA_SENSITIVITY_EPS: Final[float] = 1.0
"""VA sensitivity sign tolerance (USD)."""

# ---------------------------------------------------------------------------
# WL: $250k face, age 45 male NS, 4% flat, CSO-2017 placeholder.
# ---------------------------------------------------------------------------

WL_BENCHMARK_SP_LO: Final[float] = 30_000.0
"""Lower bound for SP-WL premium at $250k face, age 45 NS.

Industry SP-WL pricing range; depends heavily on mortality table.
Synthetic CSO 2017 placeholder rates are slightly lower than published
CSO Ultimate at age 45, so the band is widened on the lower end."""

WL_BENCHMARK_SP_HI: Final[float] = 100_000.0
"""Upper bound for SP-WL premium at $250k face, age 45 NS."""

WL_NSP_TOL: Final[float] = 1.0
"""WL net single premium vs closed form sum tolerance (USD)."""

WL_SENSITIVITY_EPS: Final[float] = 10.0
"""WL sensitivity sign tolerance (USD)."""

# ---------------------------------------------------------------------------
# UL: $250k face, $25k SP, age 45 male NS, 4% credit, 4% flat.
# ---------------------------------------------------------------------------

UL_BENCHMARK_AV_20Y_LO: Final[float] = 5_000.0
"""Lower bound for UL AV after 20 years.

After 20y of COI + expense, declared rate barely covers; lower band
allows for AV depletion approaching."""

UL_BENCHMARK_AV_20Y_HI: Final[float] = 60_000.0
"""Upper bound for UL AV after 20 years."""

UL_BENCHMARK_DEPLETION_AGE_LO: Final[int] = 70
"""Earliest plausible attained age at which UL AV depletes."""

UL_BENCHMARK_DEPLETION_AGE_HI: Final[int] = 120
"""Latest plausible attained age at which UL AV depletes (or never)."""

UL_SENSITIVITY_EPS: Final[float] = 1.0
"""UL sensitivity sign tolerance (USD)."""

# ---------------------------------------------------------------------------
# IUL: like UL but 80% pop, 10% cap, 0% floor.
# ---------------------------------------------------------------------------

IUL_BENCHMARK_AV_20Y_LO: Final[float] = 5_000.0
"""Lower bound for IUL AV after 20 years.

IUL with floor 0 dominates UL with same declared rate when index ≥ 0
cumulative. Wider lower bound to allow flat-index scenarios."""

IUL_BENCHMARK_AV_20Y_HI: Final[float] = 200_000.0
"""Upper bound for IUL AV after 20 years."""

IUL_SENSITIVITY_EPS: Final[float] = 1.0
"""IUL sensitivity sign tolerance (USD)."""

# ---------------------------------------------------------------------------
# VUL: like UL but 6% drift, 15% vol.
# ---------------------------------------------------------------------------

VUL_BENCHMARK_AV_20Y_MC_LO: Final[float] = 5_000.0
"""Lower bound for VUL E[AV(20y)] (MC mean)."""

VUL_BENCHMARK_AV_20Y_MC_HI: Final[float] = 250_000.0
"""Upper bound for VUL E[AV(20y)] (MC mean)."""

VUL_SENSITIVITY_EPS: Final[float] = 1.0
"""VUL sensitivity sign tolerance (USD)."""

# ---------------------------------------------------------------------------
# Public surface (every constant must appear in __all__ for the
# render_actuarial_benchmarks.py reflection check).
# ---------------------------------------------------------------------------
__all__ = [
    # MYGA
    "MYGA_BENCHMARK_AV_T_HI",
    "MYGA_BENCHMARK_AV_T_LO",
    "MYGA_BENCHMARK_PV_HI",
    "MYGA_BENCHMARK_PV_LO",
    "MYGA_CLOSED_FORM_AV_TOL",
    "MYGA_SENSITIVITY_EPS",
    # FIA
    "FIA_BENCHMARK_AV_T_HI",
    "FIA_BENCHMARK_AV_T_LO",
    "FIA_SENSITIVITY_EPS",
    # VA
    "VA_BENCHMARK_AV_T_FLAT_HI",
    "VA_BENCHMARK_AV_T_FLAT_LO",
    "VA_BENCHMARK_AV_T_MC_HI",
    "VA_BENCHMARK_AV_T_MC_LO",
    "VA_SENSITIVITY_EPS",
    # WL
    "WL_BENCHMARK_SP_HI",
    "WL_BENCHMARK_SP_LO",
    "WL_NSP_TOL",
    "WL_SENSITIVITY_EPS",
    # UL
    "UL_BENCHMARK_AV_20Y_HI",
    "UL_BENCHMARK_AV_20Y_LO",
    "UL_BENCHMARK_DEPLETION_AGE_HI",
    "UL_BENCHMARK_DEPLETION_AGE_LO",
    "UL_SENSITIVITY_EPS",
    # IUL
    "IUL_BENCHMARK_AV_20Y_HI",
    "IUL_BENCHMARK_AV_20Y_LO",
    "IUL_SENSITIVITY_EPS",
    # VUL
    "VUL_BENCHMARK_AV_20Y_MC_HI",
    "VUL_BENCHMARK_AV_20Y_MC_LO",
    "VUL_SENSITIVITY_EPS",
]
