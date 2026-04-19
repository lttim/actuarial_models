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


# ---------------------------------------------------------------------------
# ALM funding-ratio invariants (per product)
# ---------------------------------------------------------------------------
#
# Reading the engine (``pricing_projection.run_alm_projection_from_liability_path``):
#
#   FR_m       = AssetMV_m / (LiabPV_m + borrowing_balance_m)   (when denom > 0)
#   surplus_m  = AssetMV_m - LiabPV_m - borrowing_balance_m
#
# The "+ borrowing_balance" term is the engine's deliberate choice -- debt
# is treated as senior to liabilities, not netted against assets -- so the
# identities the property test pins are *with* the borrowing offset.
#
# Two product-agnostic laws every ALM run must obey:
#
#   1. surplus identity:        surplus == AMV - LiabPV - borrowing_balance
#   2. funding-ratio identity:  FR == AMV / (LiabPV + borrowing_balance)
#                               (where the denominator is positive)
#
# Plus a directional law tied to the asset side:
#
#   3. assets monotone in initial assets at month 0:
#        AMV_0(aum_hi) > AMV_0(aum_lo)   when aum_hi > aum_lo
#      because the initial allocation is a deterministic function of
#      ``initial_asset_market_value``, and LiabPV_0 + debt_0 is independent
#      of the asset side at issue.
#
# Hypothesis-driven so they fire across the legal demographic / curve space
# instead of only the one or two scenarios we happen to write parity tests for.
# Per-product parametrisation guards against a future engine quietly skipping
# the dispatch to ``run_alm_projection_from_pricing_result``.

_ALM_DEFAULT_ASSUMPTIONS = sp.ALMAssumptions(
    allocation=sp.alm_default_allocation_spec(),
    rebalance_band=0.10,
    rebalance_frequency_months=1,
    reinvest_rule="hold_cash",
    disinvest_rule="shortest_first",
    rebalance_policy="liquidity_only",
    liquidity_near_liquid_years=0.25,
)


def _flat_mortality(qx_value: float = 0.005) -> sp.MortalityTableQx:
    ages = np.arange(0, 121, dtype=int)
    qx = np.full_like(ages, qx_value, dtype=float)
    return sp.MortalityTableQx(ages, qx)


def _price_spia(rate: float, benefit: float, age: int) -> object:
    yc = sp.YieldCurve.from_flat_rate(rate)
    contract = sp.SPIAContract(issue_age=age, sex="male", benefit_annual=benefit)
    return sp.price_spia_single_premium(
        contract=contract,
        yield_curve=yc,
        mortality=_flat_mortality(),
        horizon_age=min(age + 25, 120),
        spread=0.0,
    )


def _price_term(rate: float, premium: float, age: int) -> object:
    yc = sp.YieldCurve.from_flat_rate(rate)
    contract = tp.TermLifeContract(
        issue_age=age,
        sex="male",
        death_benefit=250_000.0,
        monthly_premium=premium,
        term_years=20,
    )
    return tp.price_term_life_level_monthly(
        contract=contract,
        yield_curve=yc,
        mortality=_flat_mortality(),
        horizon_age=min(age + 25, 120),
        spread=0.0,
        valuation_year=None,
    )


def _price_rila(rate: float, age: int) -> object:
    yc = sp.YieldCurve.from_flat_rate(rate)
    contract = rp.RILAContract(
        issue_age=age,
        sex="male",
        participation=1.0,
        cap=0.05,
        floor=-0.05,
        rider_fee_annual=0.0,
    )
    return rp.price_rila_single_premium(
        contract=contract,
        yield_curve=yc,
        mortality=_flat_mortality(),
        horizon_age=min(age + 20, 120),
        spread=0.0,
        valuation_year=None,
        expenses=None,
    )


def _assert_alm_invariants(alm: object) -> None:
    """Pin the surplus + funding-ratio identities the ALM engine must obey.

    Folded into a helper so the per-product Hypothesis tests stay readable
    and any future invariant addition (e.g. liquidity-buffer non-negativity)
    fires for every product without copy-paste drift.
    """
    amv = alm.asset_market_value  # type: ignore[attr-defined]
    liab_pv = alm.liability_pv  # type: ignore[attr-defined]
    debt = alm.borrowing_balance  # type: ignore[attr-defined]
    surplus = alm.surplus  # type: ignore[attr-defined]
    fr = alm.funding_ratio  # type: ignore[attr-defined]

    np.testing.assert_allclose(surplus, amv - liab_pv - debt, rtol=0.0, atol=1e-6)
    denom = liab_pv + debt
    pos = denom > 1e-9
    if pos.any():
        np.testing.assert_allclose(fr[pos], amv[pos] / denom[pos], rtol=1e-9, atol=1e-9)


@given(
    rate=st.floats(min_value=0.01, max_value=0.06, allow_nan=False),
    age=st.integers(min_value=55, max_value=72),
    benefit=st.floats(min_value=20_000.0, max_value=100_000.0, allow_nan=False),
)
@_FAST_SETTINGS
def test_alm_surplus_and_funding_ratio_identities_spia(
    rate: float, age: int, benefit: float
) -> None:
    pricing = _price_spia(rate, benefit, age)
    yc = sp.YieldCurve.from_flat_rate(rate)
    alm = sp.run_alm_projection_from_pricing_result(
        pricing=pricing,
        yield_curve=yc,
        spread=0.0,
        assumptions=_ALM_DEFAULT_ASSUMPTIONS,
    )
    _assert_alm_invariants(alm)


@given(
    rate=st.floats(min_value=0.01, max_value=0.06, allow_nan=False),
    age=st.integers(min_value=30, max_value=55),
    premium=st.floats(min_value=50.0, max_value=400.0, allow_nan=False),
)
@_FAST_SETTINGS
def test_alm_surplus_and_funding_ratio_identities_term(
    rate: float, age: int, premium: float
) -> None:
    pricing = _price_term(rate, premium, age)
    yc = sp.YieldCurve.from_flat_rate(rate)
    alm = sp.run_alm_projection_from_pricing_result(
        pricing=pricing,
        yield_curve=yc,
        spread=0.0,
        assumptions=_ALM_DEFAULT_ASSUMPTIONS,
        initial_asset_market_value=500_000.0,
    )
    _assert_alm_invariants(alm)


@given(
    rate=st.floats(min_value=0.01, max_value=0.06, allow_nan=False),
    age=st.integers(min_value=45, max_value=65),
)
@_FAST_SETTINGS
def test_alm_surplus_and_funding_ratio_identities_rila(rate: float, age: int) -> None:
    pricing = _price_rila(rate, age)
    yc = sp.YieldCurve.from_flat_rate(rate)
    alm = sp.run_alm_projection_from_pricing_result(
        pricing=pricing,
        yield_curve=yc,
        spread=0.0,
        assumptions=_ALM_DEFAULT_ASSUMPTIONS,
        initial_asset_market_value=400_000.0,
    )
    _assert_alm_invariants(alm)


# Strictly-monotone-in-initial-assets at month 0. We assert the AMV
# directional law (which holds unconditionally because the initial
# allocation is a deterministic linear function of
# ``initial_asset_market_value``) and ALSO check funding ratio when the
# liability denominator is positive enough to be informative -- the
# combined check fails loudly if either the asset side stops scaling or
# the engine starts cross-coupling LiabPV_0 to the asset side.

def _assert_assets_strictly_increase(
    alm_lo: object, alm_hi: object, aum_lo: float, aum_hi: float
) -> None:
    amv_lo0 = float(alm_lo.asset_market_value[0])  # type: ignore[attr-defined]
    amv_hi0 = float(alm_hi.asset_market_value[0])  # type: ignore[attr-defined]
    assert amv_hi0 > amv_lo0, (
        f"AMV[0] must scale with initial_asset_market_value: aum {aum_lo}->{aum_hi} "
        f"gave AMV[0] {amv_lo0}->{amv_hi0}."
    )
    denom_lo = float(alm_lo.liability_pv[0] + alm_lo.borrowing_balance[0])  # type: ignore[attr-defined]
    denom_hi = float(alm_hi.liability_pv[0] + alm_hi.borrowing_balance[0])  # type: ignore[attr-defined]
    if denom_lo > 1e-3 and denom_hi > 1e-3:
        fr_lo0 = float(alm_lo.funding_ratio[0])  # type: ignore[attr-defined]
        fr_hi0 = float(alm_hi.funding_ratio[0])  # type: ignore[attr-defined]
        assert fr_hi0 > fr_lo0, (
            "FR[0] must increase when assets do (liability denom invariant w.r.t. "
            f"asset side). Got {fr_lo0} -> {fr_hi0}."
        )


@given(
    rate=st.floats(min_value=0.02, max_value=0.05, allow_nan=False),
    age=st.integers(min_value=55, max_value=70),
    benefit=st.floats(min_value=30_000.0, max_value=80_000.0, allow_nan=False),
    bump=st.floats(min_value=0.10, max_value=0.50, allow_nan=False),
)
@_FAST_SETTINGS
def test_alm_assets_strictly_increase_in_initial_assets_spia(
    rate: float, age: int, benefit: float, bump: float
) -> None:
    pricing = _price_spia(rate, benefit, age)
    yc = sp.YieldCurve.from_flat_rate(rate)
    aum_lo = float(pricing.single_premium)  # type: ignore[attr-defined]
    aum_hi = aum_lo * (1.0 + bump)
    alm_lo = sp.run_alm_projection_from_pricing_result(
        pricing=pricing, yield_curve=yc, spread=0.0,
        assumptions=_ALM_DEFAULT_ASSUMPTIONS, initial_asset_market_value=aum_lo,
    )
    alm_hi = sp.run_alm_projection_from_pricing_result(
        pricing=pricing, yield_curve=yc, spread=0.0,
        assumptions=_ALM_DEFAULT_ASSUMPTIONS, initial_asset_market_value=aum_hi,
    )
    _assert_assets_strictly_increase(alm_lo, alm_hi, aum_lo, aum_hi)


@given(
    rate=st.floats(min_value=0.02, max_value=0.05, allow_nan=False),
    age=st.integers(min_value=35, max_value=55),
    bump=st.floats(min_value=0.10, max_value=0.50, allow_nan=False),
)
@_FAST_SETTINGS
def test_alm_assets_strictly_increase_in_initial_assets_term(
    rate: float, age: int, bump: float
) -> None:
    pricing = _price_term(rate, premium=200.0, age=age)
    yc = sp.YieldCurve.from_flat_rate(rate)
    aum_lo = 500_000.0
    aum_hi = aum_lo * (1.0 + bump)
    alm_lo = sp.run_alm_projection_from_pricing_result(
        pricing=pricing, yield_curve=yc, spread=0.0,
        assumptions=_ALM_DEFAULT_ASSUMPTIONS, initial_asset_market_value=aum_lo,
    )
    alm_hi = sp.run_alm_projection_from_pricing_result(
        pricing=pricing, yield_curve=yc, spread=0.0,
        assumptions=_ALM_DEFAULT_ASSUMPTIONS, initial_asset_market_value=aum_hi,
    )
    _assert_assets_strictly_increase(alm_lo, alm_hi, aum_lo, aum_hi)


@given(
    rate=st.floats(min_value=0.02, max_value=0.05, allow_nan=False),
    age=st.integers(min_value=45, max_value=65),
    bump=st.floats(min_value=0.10, max_value=0.50, allow_nan=False),
)
@_FAST_SETTINGS
def test_alm_assets_strictly_increase_in_initial_assets_rila(
    rate: float, age: int, bump: float
) -> None:
    pricing = _price_rila(rate, age)
    yc = sp.YieldCurve.from_flat_rate(rate)
    aum_lo = 400_000.0
    aum_hi = aum_lo * (1.0 + bump)
    alm_lo = sp.run_alm_projection_from_pricing_result(
        pricing=pricing, yield_curve=yc, spread=0.0,
        assumptions=_ALM_DEFAULT_ASSUMPTIONS, initial_asset_market_value=aum_lo,
    )
    alm_hi = sp.run_alm_projection_from_pricing_result(
        pricing=pricing, yield_curve=yc, spread=0.0,
        assumptions=_ALM_DEFAULT_ASSUMPTIONS, initial_asset_market_value=aum_hi,
    )
    _assert_assets_strictly_increase(alm_lo, alm_hi, aum_lo, aum_hi)
