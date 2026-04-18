"""Property-based invariants using Hypothesis.

These complement the parity suite. Where parity asserts "Python equals Excel
for *this* scenario", these assert "Python obeys these mathematical laws for
*every* legal scenario in the input space".

Skipped if Hypothesis is not installed (it lives in requirements-dev.txt only).
"""

from __future__ import annotations

import numpy as np
import pytest

hyp = pytest.importorskip("hypothesis")
from hypothesis import HealthCheck, assume, given, settings  # noqa: E402
from hypothesis import strategies as st  # noqa: E402

import pricing_projection as sp  # noqa: E402
import rila_projection as rp  # noqa: E402
import term_projection as tp  # noqa: E402

pytestmark = pytest.mark.property

_FAST_SETTINGS = settings(
    max_examples=25,
    deadline=2_000,
    suppress_health_check=[HealthCheck.too_slow, HealthCheck.large_base_example],
)


# ---------------------------------------------------------------------------
# SPIA invariants
# ---------------------------------------------------------------------------


@given(
    issue_age=st.integers(min_value=40, max_value=85),
    benefit_annual=st.floats(min_value=1_000.0, max_value=200_000.0, allow_nan=False),
    rate=st.floats(min_value=0.005, max_value=0.10, allow_nan=False),
)
@_FAST_SETTINGS
def test_spia_single_premium_is_positive_when_benefit_positive(
    issue_age: int, benefit_annual: float, rate: float
) -> None:
    """A positive benefit at a non-degenerate rate must produce a positive SP."""
    contract = sp.SPIAContract(issue_age=issue_age, sex="male", benefit_annual=benefit_annual)
    yc = sp.YieldCurve.from_flat_rate(rate)
    ages = np.arange(0, 121, dtype=int)
    qx = np.clip(0.005 + ages * 5e-5, 1e-6, 0.4)
    mort = sp.MortalityTableQx(ages, qx)
    res = sp.price_spia_single_premium(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=min(issue_age + 30, 120),
        spread=0.0,
    )
    assert res.single_premium > 0
    assert np.isfinite(res.single_premium)


@given(
    benefit_annual=st.floats(min_value=10_000.0, max_value=120_000.0, allow_nan=False),
    rate=st.floats(min_value=0.01, max_value=0.08, allow_nan=False),
)
@_FAST_SETTINGS
def test_spia_higher_rate_produces_lower_single_premium(benefit_annual: float, rate: float) -> None:
    """At higher discount rate, the SP for the same benefit must be strictly lower."""
    ages = np.arange(0, 121, dtype=int)
    qx = np.clip(0.005 + ages * 5e-5, 1e-6, 0.4)
    mort = sp.MortalityTableQx(ages, qx)
    contract = sp.SPIAContract(issue_age=65, sex="male", benefit_annual=benefit_annual)
    res_lo = sp.price_spia_single_premium(
        contract=contract,
        yield_curve=sp.YieldCurve.from_flat_rate(rate),
        mortality=mort,
        horizon_age=95,
        spread=0.0,
    )
    res_hi = sp.price_spia_single_premium(
        contract=contract,
        yield_curve=sp.YieldCurve.from_flat_rate(rate + 0.01),
        mortality=mort,
        horizon_age=95,
        spread=0.0,
    )
    assert res_hi.single_premium < res_lo.single_premium


# ---------------------------------------------------------------------------
# Term Life invariants
# ---------------------------------------------------------------------------


@given(
    issue_age=st.integers(min_value=25, max_value=70),
    death_benefit=st.floats(min_value=10_000.0, max_value=2_000_000.0, allow_nan=False),
    monthly_premium=st.floats(min_value=10.0, max_value=2_000.0, allow_nan=False),
    term_years=st.integers(min_value=10, max_value=30),
)
@_FAST_SETTINGS
def test_term_zero_qx_means_zero_claims(
    issue_age: int, death_benefit: float, monthly_premium: float, term_years: int
) -> None:
    """If q_x is zero everywhere, expected claim cashflow must be exactly zero."""
    contract = tp.TermLifeContract(
        issue_age=issue_age,
        sex="male",
        death_benefit=death_benefit,
        monthly_premium=monthly_premium,
        term_years=term_years,
    )
    ages = np.arange(0, 121, dtype=int)
    qx = np.zeros_like(ages, dtype=float)
    mort = sp.MortalityTableQx(ages, qx)
    yc = sp.YieldCurve.from_flat_rate(0.04)
    res = tp.price_term_life_level_monthly(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=min(issue_age + term_years + 5, 120),
        spread=0.0,
        valuation_year=None,
    )
    assert np.allclose(res.expected_claim_cashflows, 0.0)


# ---------------------------------------------------------------------------
# RILA invariants
# ---------------------------------------------------------------------------


@given(
    raw=st.floats(min_value=-0.5, max_value=0.5, allow_nan=False),
    cap=st.floats(min_value=0.02, max_value=0.20, allow_nan=False),
    floor=st.floats(min_value=-0.20, max_value=-0.001, allow_nan=False),
    participation=st.floats(min_value=0.5, max_value=1.5, allow_nan=False),
)
@_FAST_SETTINGS
def test_rila_credited_return_within_floor_and_cap(
    raw: float, cap: float, floor: float, participation: float
) -> None:
    """Credited return is bounded by [floor, cap] for any raw input."""
    assume(cap > floor)
    cr = rp.segment_credited_return(raw=raw, participation=participation, cap=cap, floor=floor)
    assert floor - 1e-12 <= cr <= cap + 1e-12


@given(
    raw=st.floats(min_value=0.0, max_value=0.30, allow_nan=False),
    cap=st.floats(min_value=0.02, max_value=0.10, allow_nan=False),
    floor=st.floats(min_value=-0.20, max_value=-0.01, allow_nan=False),
)
@_FAST_SETTINGS
def test_rila_credited_return_monotone_in_raw_when_participation_positive(
    raw: float, cap: float, floor: float
) -> None:
    """Increasing raw return cannot decrease credited return (monotonicity)."""
    cr_lo = rp.segment_credited_return(raw=raw, participation=1.0, cap=cap, floor=floor)
    cr_hi = rp.segment_credited_return(raw=raw + 0.01, participation=1.0, cap=cap, floor=floor)
    assert cr_hi >= cr_lo - 1e-12
