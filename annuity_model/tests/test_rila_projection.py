from __future__ import annotations

import numpy as np
import pytest

import pricing_projection as sp
import rila_projection as rp

pytestmark = pytest.mark.product_rila

# Non-zero monthly expense so implicit single-premium pricing is well-posed
# (all-zero expenses degenerate to premium 0 and are rejected by the engine).
_RILA_EX = sp.ExpenseAssumptions(0.0, 0.0, 25.0)


def test_all_zero_expense_assumptions_rejected():
    contract = rp.RILAContract(
        issue_age=65,
        sex="male",
        participation=0.8,
        cap=0.1,
        floor=0.0,
        rider_fee_annual=0.01,
    )
    yc = sp.YieldCurve.from_flat_rate(0.04)
    ages = np.arange(0, 121, dtype=int)
    qx = np.full_like(ages, 0.02, dtype=float)
    mort = sp.MortalityTableQx(ages, qx)
    with pytest.raises(ValueError, match="non-positive"):
        rp.price_rila_single_premium(
            contract=contract,
            yield_curve=yc,
            mortality=mort,
            horizon_age=80,
            spread=0.0,
            valuation_year=None,
            expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
            index_scenario_csv_path=None,
            expense_annual_inflation=0.0,
        )


def test_segment_credited_return_respects_cap_floor():
    assert rp.segment_credited_return(
        raw=0.2, participation=1.0, cap=0.10, floor=0.0
    ) == pytest.approx(0.10)
    assert rp.segment_credited_return(
        raw=-0.5, participation=1.0, cap=0.10, floor=-0.05
    ) == pytest.approx(-0.05)
    assert rp.segment_credited_return(
        raw=0.05, participation=0.8, cap=0.10, floor=0.0
    ) == pytest.approx(0.04)


def test_flat_index_zero_crediting_and_positive_premium():
    """Flat index + zero rider fee implies zero segment crediting; premium is expense-driven."""
    contract = rp.RILAContract(
        issue_age=65,
        sex="male",
        participation=1.0,
        cap=0.10,
        floor=0.0,
        rider_fee_annual=0.0,
    )
    yc = sp.YieldCurve.from_flat_rate(0.04)
    ages = np.arange(0, 121, dtype=int)
    qx = np.clip(0.01 + ages * 1e-5, 1e-6, 0.35)
    mort = sp.MortalityTableQx(ages, qx)
    ex = _RILA_EX
    res = rp.price_rila_single_premium(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=80,
        spread=0.0,
        valuation_year=None,
        expenses=ex,
        index_scenario_csv_path=None,
        expense_annual_inflation=0.0,
    )
    res_dup = rp.price_rila_single_premium(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=80,
        spread=0.0,
        valuation_year=None,
        expenses=ex,
        index_scenario_csv_path=None,
        expense_annual_inflation=0.0,
    )
    assert res.single_premium > 0.0
    np.testing.assert_allclose(res.single_premium, res_dup.single_premium, rtol=0, atol=1e-6)
    assert np.allclose(res.segment_credited_rate, 0.0)


def test_liability_path_alm_runs():
    contract = rp.RILAContract(
        issue_age=60,
        sex="male",
        participation=0.9,
        cap=0.08,
        floor=0.0,
        rider_fee_annual=0.005,
    )
    yc = sp.YieldCurve.from_flat_rate(0.03)
    ages = np.arange(0, 121, dtype=int)
    qx = np.full_like(ages, 0.02, dtype=float)
    mort = sp.MortalityTableQx(ages, qx)
    ex = _RILA_EX
    res = rp.price_rila_single_premium(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=75,
        spread=0.0,
        valuation_year=None,
        expenses=ex,
        index_scenario_csv_path=None,
        expense_annual_inflation=0.0,
    )
    path = rp.liability_path_from_rila_projection(res)
    asm = sp.ALMAssumptions(
        allocation=sp.alm_default_allocation_spec(),
        rebalance_band=0.10,
        rebalance_frequency_months=1,
        reinvest_rule="hold_cash",
        disinvest_rule="shortest_first",
        rebalance_policy="liquidity_only",
        liquidity_near_liquid_years=0.25,
    )
    alm = sp.run_alm_projection_from_liability_path(
        liability_path=path,
        yield_curve=yc,
        spread=0.0,
        assumptions=asm,
        initial_asset_market_value=float(res.single_premium) + 50_000.0,
    )
    assert alm.surplus.shape == res.expected_total_cashflows.shape
    assert np.isfinite(alm.surplus).all()


def test_monte_carlo_shape():
    contract = rp.RILAContract(
        issue_age=70,
        sex="female",
        participation=1.0,
        cap=0.12,
        floor=0.0,
        rider_fee_annual=0.01,
    )
    yc = sp.YieldCurve.from_flat_rate(0.04)
    ages = np.arange(0, 121, dtype=int)
    qx = np.full_like(ages, 0.03, dtype=float)
    mort = sp.MortalityTableQx(ages, qx)
    mc = rp.price_rila_single_premium_monte_carlo(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=85,
        spread=0.0,
        valuation_year=None,
        expenses=_RILA_EX,
        n_sims=20,
        annual_drift=0.05,
        annual_vol=0.12,
        seed=123,
        s0=100.0,
    )
    assert mc.single_premium.shape == (20,)
    assert np.isfinite(mc.premium_mean)


def test_pricing_infeasible_raises_with_loading_details():
    """Aggressive crediting can push K + premium_expense_rate to 1 or above."""
    contract = rp.RILAContract(
        issue_age=65,
        sex="male",
        participation=5.0,
        cap=2.0,
        floor=0.0,
        rider_fee_annual=0.01,
    )
    yc = sp.YieldCurve.from_flat_rate(0.04)
    ages = np.arange(0, 121, dtype=int)
    qx = np.clip(0.005 + ages * 1e-5, 1e-6, 0.4)
    mort = sp.MortalityTableQx(ages, qx)
    ex = sp.ExpenseAssumptions.load_from_csv(sp.DEFAULT_EXPENSES_CSV)
    levels = np.asarray(
        sp.load_index_scenario_monthly_csv(sp.DEFAULT_SP500_SCENARIO_CSV, n_months=540)[1],
        dtype=float,
    )
    with pytest.raises(rp.RILAPricingInfeasibleError) as excinfo:
        rp.price_rila_single_premium(
            contract=contract,
            yield_curve=yc,
            mortality=mort,
            horizon_age=110,
            spread=0.0,
            valuation_year=None,
            expenses=ex,
            index_s0=100.0,
            index_levels_payment=levels,
            expense_annual_inflation=0.0,
        )
    err = excinfo.value
    assert "K=" in str(err)
    assert err.k_loading + err.premium_expense_rate >= 1.0 - 1e-12


def test_monte_carlo_skip_policy_records_infeasible_paths():
    """Default UI parameters can trip per-path infeasibility on long horizons; MC must keep going."""
    contract = rp.RILAContract(
        issue_age=65,
        sex="male",
        participation=1.0,
        cap=0.10,
        floor=0.0,
        rider_fee_annual=0.01,
    )
    yc = sp.YieldCurve.from_flat_rate(0.04)
    ages = np.arange(0, 121, dtype=int)
    qx = np.clip(0.005 + ages * 1e-5, 1e-6, 0.4)
    mort = sp.MortalityTableQx(ages, qx)
    mc = rp.price_rila_single_premium_monte_carlo(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=110,
        spread=0.0,
        valuation_year=None,
        expenses=_RILA_EX,
        n_sims=100,
        annual_drift=0.06,
        annual_vol=0.15,
        seed=42,
        s0=100.0,
    )
    assert mc.n_sims == 100
    assert mc.n_feasible + mc.n_infeasible == 100
    assert mc.n_feasible >= 1
    nan_mask = np.isnan(mc.single_premium)
    assert int(nan_mask.sum()) == mc.n_infeasible
    if mc.n_infeasible > 0:
        assert mc.infeasible_max_loading >= 1.0 - 1e-12
    assert np.isfinite(mc.premium_mean)
    assert np.isfinite(mc.premium_median)


def test_monte_carlo_raise_policy_preserves_legacy_behavior():
    contract = rp.RILAContract(
        issue_age=65,
        sex="male",
        participation=5.0,
        cap=2.0,
        floor=0.0,
        rider_fee_annual=0.01,
    )
    yc = sp.YieldCurve.from_flat_rate(0.04)
    ages = np.arange(0, 121, dtype=int)
    qx = np.clip(0.005 + ages * 1e-5, 1e-6, 0.4)
    mort = sp.MortalityTableQx(ages, qx)
    with pytest.raises(rp.RILAPricingInfeasibleError):
        rp.price_rila_single_premium_monte_carlo(
            contract=contract,
            yield_curve=yc,
            mortality=mort,
            horizon_age=110,
            spread=0.0,
            valuation_year=None,
            expenses=_RILA_EX,
            n_sims=10,
            annual_drift=0.06,
            annual_vol=0.15,
            seed=7,
            s0=100.0,
            infeasible_path_policy="raise",
        )


def test_monte_carlo_benign_params_have_no_infeasible_paths():
    """A benign contract under low drift should price cleanly with zero infeasible paths."""
    contract = rp.RILAContract(
        issue_age=70,
        sex="female",
        participation=1.0,
        cap=0.10,
        floor=0.0,
        rider_fee_annual=0.01,
    )
    yc = sp.YieldCurve.from_flat_rate(0.04)
    ages = np.arange(0, 121, dtype=int)
    qx = np.full_like(ages, 0.03, dtype=float)
    mort = sp.MortalityTableQx(ages, qx)
    mc = rp.price_rila_single_premium_monte_carlo(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=85,
        spread=0.0,
        valuation_year=None,
        expenses=_RILA_EX,
        n_sims=40,
        annual_drift=0.04,
        annual_vol=0.12,
        seed=11,
        s0=100.0,
    )
    assert mc.n_infeasible == 0
    assert mc.n_feasible == mc.n_sims
    assert np.all(np.isfinite(mc.single_premium))


def test_monte_carlo_all_paths_infeasible_raises():
    """If 100% of paths fail, MC raises with the worst observed loading."""
    contract = rp.RILAContract(
        issue_age=65,
        sex="male",
        participation=10.0,
        cap=10.0,
        floor=0.0,
        rider_fee_annual=0.0,
    )
    yc = sp.YieldCurve.from_flat_rate(0.04)
    ages = np.arange(0, 121, dtype=int)
    qx = np.full_like(ages, 0.05, dtype=float)
    mort = sp.MortalityTableQx(ages, qx)
    with pytest.raises(rp.RILAPricingInfeasibleError) as excinfo:
        rp.price_rila_single_premium_monte_carlo(
            contract=contract,
            yield_curve=yc,
            mortality=mort,
            horizon_age=110,
            spread=0.0,
            valuation_year=None,
            expenses=_RILA_EX,
            n_sims=8,
            annual_drift=0.30,
            annual_vol=0.05,
            seed=1,
            s0=100.0,
        )
    assert "All 8 Monte Carlo paths" in str(excinfo.value)
