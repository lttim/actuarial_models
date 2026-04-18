"""parity_constants -- single source of truth for parity tolerances.

Every parity tolerance, Excel-formula epsilon, and numerical guard used by the
test suite or model contract documentation is exported from this module. The
goal is that the two parity contracts under :mod:`annuity_model.docs` and the
release checklist render their tolerance tables from this file (see
``scripts/render_parity_contract.py``) so a single value change is reflected
across docs, tests, and CI in one commit.

Naming convention
-----------------
* ``TOL_*``                -- python-vs-excel comparison tolerance
* ``EXCEL_*``              -- guard / epsilon used inside Excel formulas
* ``MODELCHECK_TOL``       -- per-cell ModelCheck tolerance (zero by contract)
* ``RILA_*``               -- RILA-specific PV / AV tolerances

Anything new added to the test suite MUST be added here first.
"""

from __future__ import annotations

from typing import Final

TOL_DOLLAR: Final[float] = 1e-4
"""Cash, face, market value, surplus dollar tolerance (USD)."""

TOL_TENOR: Final[float] = 1e-6
"""Remaining-tenor tolerance (years)."""

TOL_DF: Final[float] = 1e-10
"""Discount factor tolerance (dimensionless)."""

MODELCHECK_TOL: Final[float] = 0.0
"""ModelCheck per-cell snapshot tolerance: must be exact."""

EXCEL_DISINVEST_EPSILON: Final[float] = 1e-9
"""Per-bucket epsilon (Excel) -- ``(k+1) * EXCEL_DISINVEST_EPSILON``.

Mirrors the Python pricing engine's ``np.arange(n) * 1e-9`` epsilon.
The +1 offset on the Excel side ensures the k=0 bucket has a non-zero
epsilon for the ``tmin`` comparison.
"""

EXCEL_DISINVEST_THRESHOLD: Final[float] = 5e-10
"""Tie-break threshold for Excel disinvestment (half the inter-bucket interval).

A bucket is treated as "matching ``tmin``" iff
``abs(t_rem[k] - tmin) < EXCEL_DISINVEST_THRESHOLD``.
"""

DISINVEST_TIE_BREAK_EPS: Final[float] = 1e-9
"""Python epsilon added to argsort key for stable tie-breaking by index."""

T_REM_RESET_EPS: Final[float] = 1e-9
"""Threshold below which a bucket's face / tenor is treated as "depleted".

Both ``f_pm <= 1e-9`` and ``t_pm <= 1e-9`` must hold to gate the
post-maturity ``t_rem`` reset.
"""

REINVEST_GAP_EPS: Final[float] = 1e-9
"""Minimum positive gap required before pro-rata reinvestment fires."""

REINVEST_XSR_EPS: Final[float] = 1e-6
"""Minimum excess-cash threshold for reinvestment to fire."""

# RILA-specific tolerances ---------------------------------------------------

RILA_PV_TOL: Final[float] = 1e-4
"""RILA PV / single-premium dollar tolerance (matches docs/rila_parity_contract.md)."""

RILA_AV_TOL: Final[float] = 1e-6
"""RILA account-value tolerance for crediting paths."""

# Term-life tolerances -------------------------------------------------------

TERM_MODELCHECK_TOL: Final[float] = 1e-9
"""Term Life ModelCheck cell tolerance (claims/premium/net)."""

# Pricing UI / what-if -------------------------------------------------------

WHATIF_NO_OP_TOL: Final[float] = 1e-9
"""When all what-if shocks are zero, output must match base within this tolerance."""

__all__ = [
    "DISINVEST_TIE_BREAK_EPS",
    "EXCEL_DISINVEST_EPSILON",
    "EXCEL_DISINVEST_THRESHOLD",
    "MODELCHECK_TOL",
    "REINVEST_GAP_EPS",
    "REINVEST_XSR_EPS",
    "RILA_AV_TOL",
    "RILA_PV_TOL",
    "TERM_MODELCHECK_TOL",
    "TOL_DF",
    "TOL_DOLLAR",
    "TOL_TENOR",
    "T_REM_RESET_EPS",
    "WHATIF_NO_OP_TOL",
]
