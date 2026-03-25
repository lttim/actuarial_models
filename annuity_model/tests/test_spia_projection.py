"""Unit tests for spia_projection (SPIA pricing scaffold)."""

from __future__ import annotations

import math
from pathlib import Path

import numpy as np
import pandas as pd
import pytest

import spia_projection as sp


# --- Monte Carlo first principles ---


def _minimal_mortality_mc():
    """Minimal mortality helper used across Monte Carlo tests."""
    ages = np.arange(0, 121, dtype=int)
    qx = np.clip(0.005 + ages * 1e-5, 1e-6, 0.4)
    return sp.MortalityTableQx(ages, qx)


def _mc_contract_and_setup():
    """Standard contract/yc/expenses for MC tests."""
    contract = sp.SPIAContract(issue_age=65, sex="male", benefit_annual=100_000.0)
    yc = sp.YieldCurve.from_flat_rate(0.03)
    ex = sp.ExpenseAssumptions(0.0, 0.0, 0.0)
    return contract, yc, _minimal_mortality_mc(), ex


@pytest.mark.parametrize(
    "kwargs,match",
    [
        ({"n_sims": 0, "n_months": 12, "s0": 100.0, "annual_vol": 0.15}, "n_sims"),
        ({"n_sims": 5, "n_months": 0, "s0": 100.0, "annual_vol": 0.15}, "n_months"),
        ({"n_sims": 5, "n_months": 12, "s0": 0.0, "annual_vol": 0.15}, "s0"),
        ({"n_sims": 5, "n_months": 12, "s0": 100.0, "annual_vol": -0.01}, "annual_vol"),
    ],
)
def test_simulate_index_levels_gbm_rejects_invalid_inputs(kwargs, match):
    """GBM simulator must reject n_sims<1, n_months<1, s0<=0, or annual_vol<0 with a clear ValueError."""
    with pytest.raises(ValueError, match=match):
        sp.simulate_index_levels_gbm(annual_drift=0.06, **kwargs)


def test_simulate_index_levels_gbm_output_shape():
    """GBM output has shape (n_sims, n_months+1) with column 0 equal to s0 for every simulation."""
    n_sims, n_months, s0 = 8, 24, 150.0
    paths = sp.simulate_index_levels_gbm(n_sims=n_sims, n_months=n_months, s0=s0, seed=0)
    assert paths.shape == (n_sims, n_months + 1)
    np.testing.assert_array_equal(paths[:, 0], np.full(n_sims, s0))


def test_simulate_index_levels_gbm_all_levels_strictly_positive():
    """GBM paths must remain strictly positive for all months and simulations (log-normal property)."""
    paths = sp.simulate_index_levels_gbm(
        n_sims=200, n_months=120, s0=100.0, annual_drift=0.0, annual_vol=0.3, seed=99
    )
    assert np.all(paths > 0.0)


def test_simulate_index_levels_gbm_one_step_logreturn_moments_match_theory():
    """One-month log-return sample mean and variance must match GBM theoretical moments within Monte Carlo error.

    For GBM with drift mu and volatility sigma:
      E[log(S_1/S_0)] = (mu - 0.5*sigma^2) * dt
      Var[log(S_1/S_0)] = sigma^2 * dt
    where dt = 1/12 (monthly).
    Tolerance: 5 Monte Carlo standard errors for each statistic.
    """
    n_sims = 100_000
    mu, sigma, s0, dt = 0.08, 0.20, 100.0, 1.0 / 12.0
    paths = sp.simulate_index_levels_gbm(
        n_sims=n_sims, n_months=1, s0=s0, annual_drift=mu, annual_vol=sigma, seed=42
    )
    log_ret = np.log(paths[:, 1] / paths[:, 0])

    theoretical_mean = (mu - 0.5 * sigma * sigma) * dt
    theoretical_var = sigma * sigma * dt

    sample_mean = float(np.mean(log_ret))
    sample_var = float(np.var(log_ret, ddof=1))

    # Monte Carlo SE for mean = theoretical_std / sqrt(n_sims)
    se_mean = math.sqrt(theoretical_var / n_sims)
    # Monte Carlo SE for variance ≈ sigma_lr^2 * sqrt(2/(n-1))
    se_var = theoretical_var * math.sqrt(2.0 / (n_sims - 1))

    assert abs(sample_mean - theoretical_mean) < 5.0 * se_mean, (
        f"Log-return mean {sample_mean:.6f} too far from theory {theoretical_mean:.6f}"
    )
    assert abs(sample_var - theoretical_var) < 5.0 * se_var, (
        f"Log-return variance {sample_var:.6f} too far from theory {theoretical_var:.6f}"
    )


def test_simulate_index_levels_gbm_logreturn_variance_scales_with_time():
    """Log-return variance must scale linearly with the number of monthly steps (Brownian motion property).

    Var[log(S_k/S_0)] = k * sigma^2 * dt, so Var at month k should be k times Var at month 1.
    """
    n_sims = 80_000
    sigma = 0.20
    paths = sp.simulate_index_levels_gbm(
        n_sims=n_sims, n_months=12, s0=100.0, annual_drift=0.0, annual_vol=sigma, seed=7
    )
    # Log-return from month 0 to month k
    log_ret_1 = np.log(paths[:, 1] / paths[:, 0])
    log_ret_6 = np.log(paths[:, 6] / paths[:, 0])
    var_1 = float(np.var(log_ret_1, ddof=1))
    var_6 = float(np.var(log_ret_6, ddof=1))
    ratio = var_6 / var_1
    # Expect ratio ≈ 6; allow 5% relative tolerance for Monte Carlo noise
    assert abs(ratio - 6.0) / 6.0 < 0.05, f"Variance ratio month-6/month-1 is {ratio:.4f}, expected ~6.0"


def test_simulate_index_levels_gbm_is_seed_reproducible():
    """Monte Carlo path generation should be exactly reproducible for a fixed seed."""
    p1 = sp.simulate_index_levels_gbm(n_sims=4, n_months=12, seed=123, annual_drift=0.05, annual_vol=0.2, s0=100.0)
    p2 = sp.simulate_index_levels_gbm(n_sims=4, n_months=12, seed=123, annual_drift=0.05, annual_vol=0.2, s0=100.0)
    np.testing.assert_allclose(p1, p2, rtol=0.0, atol=0.0)


def test_simulate_index_levels_gbm_different_seeds_produce_different_paths():
    """Two different seeds must produce statistically distinct path arrays (not bit-identical)."""
    p1 = sp.simulate_index_levels_gbm(n_sims=10, n_months=12, seed=1, annual_drift=0.05, annual_vol=0.2, s0=100.0)
    p2 = sp.simulate_index_levels_gbm(n_sims=10, n_months=12, seed=2, annual_drift=0.05, annual_vol=0.2, s0=100.0)
    assert not np.allclose(p1, p2), "Paths from different seeds should not be identical"


def test_simulate_index_levels_gbm_zero_vol_has_common_deterministic_path():
    """With zero volatility, all simulations follow the same deterministic drift path."""
    paths = sp.simulate_index_levels_gbm(n_sims=3, n_months=6, seed=7, annual_drift=0.12, annual_vol=0.0, s0=100.0)
    np.testing.assert_allclose(paths[0], paths[1], rtol=1e-12, atol=1e-12)
    np.testing.assert_allclose(paths[1], paths[2], rtol=1e-12, atol=1e-12)
    assert paths[0, 1] == pytest.approx(100.0 * math.exp(0.12 / 12.0), rel=1e-12)


def test_price_spia_monte_carlo_quantiles_are_ordered():
    """MC output must satisfy p05 <= median <= p95 and all summary metrics must be finite."""
    contract, yc, mort, ex = _mc_contract_and_setup()
    mc = sp.price_spia_single_premium_monte_carlo(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=80,
        expenses=ex,
        n_sims=500,
        annual_drift=0.06,
        annual_vol=0.18,
        seed=5,
        s0=100.0,
        expense_annual_inflation=0.0,
    )
    assert mc.premium_p05 <= mc.premium_median <= mc.premium_p95, (
        f"Quantile order violated: p05={mc.premium_p05:.2f} median={mc.premium_median:.2f} p95={mc.premium_p95:.2f}"
    )
    for attr in ("premium_mean", "premium_median", "premium_p05", "premium_p95", "pv_benefit_mean", "pv_total_mean"):
        val = getattr(mc, attr)
        assert math.isfinite(val), f"{attr}={val} is not finite"


def test_price_spia_monte_carlo_seed_reproducible_full_result():
    """Identical seed must yield bitwise-identical premium arrays and all summary statistics."""
    contract, yc, mort, ex = _mc_contract_and_setup()
    kwargs = dict(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=80,
        expenses=ex,
        n_sims=200,
        annual_drift=0.06,
        annual_vol=0.15,
        seed=99,
        s0=100.0,
        expense_annual_inflation=0.0,
    )
    mc1 = sp.price_spia_single_premium_monte_carlo(**kwargs)
    mc2 = sp.price_spia_single_premium_monte_carlo(**kwargs)
    np.testing.assert_array_equal(mc1.single_premium, mc2.single_premium)
    assert mc1.premium_mean == mc2.premium_mean
    assert mc1.premium_p05 == mc2.premium_p05
    assert mc1.premium_p95 == mc2.premium_p95


def test_price_spia_monte_carlo_different_seed_changes_results():
    """Different random seeds must produce materially different premium distributions."""
    contract, yc, mort, ex = _mc_contract_and_setup()
    common = dict(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=80,
        expenses=ex,
        n_sims=300,
        annual_drift=0.06,
        annual_vol=0.20,
        s0=100.0,
        expense_annual_inflation=0.0,
    )
    mc1 = sp.price_spia_single_premium_monte_carlo(**common, seed=11)
    mc2 = sp.price_spia_single_premium_monte_carlo(**common, seed=99)
    assert not np.array_equal(mc1.single_premium, mc2.single_premium), (
        "Premium arrays from different seeds should not be identical"
    )


def test_price_spia_monte_carlo_matches_manual_path_loop():
    """MC wrapper output must match a manually written path-by-path pricing loop exactly."""
    contract, yc, mort, ex = _mc_contract_and_setup()
    n_sims, n_months = 50, int(round((80 - contract.issue_age) * 12))
    drift, vol, s0, seed = 0.05, 0.12, 100.0, 77

    # Build paths identically to the wrapper
    idx_paths = sp.simulate_index_levels_gbm(
        n_sims=n_sims, n_months=n_months, s0=s0, annual_drift=drift, annual_vol=vol, seed=seed
    )

    # Reprice manually path by path
    manual_premiums = np.zeros(n_sims, dtype=float)
    for i in range(n_sims):
        res_i = sp.price_spia_single_premium(
            contract=contract,
            yield_curve=yc,
            mortality=mort,
            horizon_age=80,
            expenses=ex,
            index_s0=float(idx_paths[i, 0]),
            index_levels_payment=idx_paths[i, 1:],
            expense_annual_inflation=0.0,
        )
        manual_premiums[i] = float(res_i.single_premium)

    mc = sp.price_spia_single_premium_monte_carlo(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=80,
        expenses=ex,
        n_sims=n_sims,
        annual_drift=drift,
        annual_vol=vol,
        seed=seed,
        s0=s0,
        expense_annual_inflation=0.0,
    )
    np.testing.assert_allclose(mc.single_premium, manual_premiums, rtol=1e-12, atol=0.0)
    assert mc.premium_mean == pytest.approx(float(np.mean(manual_premiums)), rel=1e-12)


def test_price_spia_monte_carlo_higher_vol_increases_distribution_width():
    """A higher index volatility must produce a wider premium distribution (p95 - p05 spread increases).

    This validates that the model correctly translates index uncertainty into pricing uncertainty.
    """
    contract, yc, mort, ex = _mc_contract_and_setup()
    common = dict(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=80,
        expenses=ex,
        n_sims=1000,
        annual_drift=0.05,
        seed=42,
        s0=100.0,
        expense_annual_inflation=0.0,
    )
    mc_low_vol = sp.price_spia_single_premium_monte_carlo(**common, annual_vol=0.05)
    mc_high_vol = sp.price_spia_single_premium_monte_carlo(**common, annual_vol=0.35)

    spread_low = mc_low_vol.premium_p95 - mc_low_vol.premium_p05
    spread_high = mc_high_vol.premium_p95 - mc_high_vol.premium_p05
    assert spread_high > spread_low, (
        f"Higher vol should widen p95-p05 spread: low_vol_spread={spread_low:.2f}, high_vol_spread={spread_high:.2f}"
    )


def test_price_spia_monte_carlo_zero_vol_matches_deterministic_mean():
    """At zero vol, MC premium distribution collapses to the deterministic premium value."""
    contract = sp.SPIAContract(issue_age=65, sex="male", benefit_annual=100_000.0)
    yc = sp.YieldCurve.from_flat_rate(0.03)
    mort = _minimal_mortality_mc()
    ex = sp.ExpenseAssumptions(0.0, 0.0, 0.0)

    n_months = int(round((80 - contract.issue_age) * 12))
    deterministic_levels = sp.simulate_index_levels_gbm(
        n_sims=1,
        n_months=n_months,
        s0=100.0,
        annual_drift=0.0,
        annual_vol=0.0,
        seed=1,
    )[0]
    det = sp.price_spia_single_premium(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=80,
        expenses=ex,
        index_s0=float(deterministic_levels[0]),
        index_levels_payment=deterministic_levels[1:],
        expense_annual_inflation=0.0,
    )
    mc = sp.price_spia_single_premium_monte_carlo(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=80,
        expenses=ex,
        n_sims=200,
        annual_drift=0.0,
        annual_vol=0.0,
        seed=1,
        s0=100.0,
        expense_annual_inflation=0.0,
    )
    assert mc.premium_mean == pytest.approx(det.single_premium, rel=1e-10)
    assert mc.premium_p05 == pytest.approx(det.single_premium, rel=1e-10)
    assert mc.premium_p95 == pytest.approx(det.single_premium, rel=1e-10)


# --- YieldCurve ---


def test_yield_curve_from_flat_rate_discount_factors():
    """Confirms a flat zero curve discounts as DF(t) = exp(-z×t), i.e. continuous compounding on the curve."""
    z = 0.04
    yc = sp.YieldCurve.from_flat_rate(z)
    t = np.array([0.0, 0.5, 1.0, 2.0])
    df = yc.discount_factors(t)
    expected = np.exp(-z * t)
    np.testing.assert_allclose(df, expected, rtol=1e-12)


def test_yield_curve_discount_factors_with_spread():
    """Checks that a credit/spread add-on is added to the zero rate when building discount factors."""
    yc = sp.YieldCurve.from_flat_rate(0.03)
    t = np.array([1.0])
    df = yc.discount_factors(t, spread=0.01)
    assert df[0] == pytest.approx(math.exp(-0.04), rel=1e-12)


def test_yield_curve_discount_factors_rejects_non_continuous():
    """Ensures the model refuses non-continuous compounding (only continuous is implemented)."""
    yc = sp.YieldCurve.from_flat_rate(0.03)
    with pytest.raises(ValueError, match="continuous"):
        yc.discount_factors(np.array([1.0]), compounding="annual")  # type: ignore[arg-type]


def test_yield_curve_discount_factors_rejects_non_1d():
    """Ensures discount factors are only computed for a 1D list of times (guards misuse)."""
    yc = sp.YieldCurve.from_flat_rate(0.03)
    with pytest.raises(ValueError, match="1D"):
        yc.discount_factors(np.array([[1.0, 2.0]]))


def test_yield_curve_extrapolation_below_and_above_nodes():
    """Verifies interpolation between nodes and flat zero-rate extrapolation before the first and after the last node."""
    mats = np.array([1.0, 3.0])
    zeros = np.array([0.02, 0.05])
    yc = sp.YieldCurve(mats, zeros)
    spread = 0.005
    t = np.array([0.25, 2.0, 10.0])
    df = yc.discount_factors(t, spread=spread)
    assert df[0] == pytest.approx(math.exp(-(0.02 + spread) * 0.25), rel=1e-9)
    assert df[2] == pytest.approx(math.exp(-(0.05 + spread) * 10.0), rel=1e-9)
    # Midpoint: log-linear interpolation on DF between nodes
    log_df_1 = math.log(math.exp(-(0.02 + spread) * 1.0))
    log_df_3 = math.log(math.exp(-(0.05 + spread) * 3.0))
    log_df_2 = np.interp(2.0, mats, np.array([log_df_1, log_df_3]))
    assert df[1] == pytest.approx(math.exp(log_df_2), rel=1e-9)


def test_yield_curve_load_zero_curve_csv_sorts_by_maturity(tmp_path):
    """Checks loading a zero curve from CSV sorts maturities ascending (so the curve is well ordered)."""
    p = tmp_path / "z.csv"
    p.write_text("maturity_years,zero_rate\n2.0,0.05\n0.5,0.01\n", encoding="utf-8")
    yc = sp.YieldCurve.load_zero_curve_csv(str(p))
    np.testing.assert_array_equal(yc.maturities_years, np.array([0.5, 2.0]))
    np.testing.assert_array_equal(yc.zero_rates, np.array([0.01, 0.05]))


# --- bootstrap_zero_rates_from_par_yields ---


def test_bootstrap_rejects_bad_coupon_freq():
    """Ensures par-yield bootstrapping rejects invalid coupon frequency (must be positive)."""
    with pytest.raises(ValueError, match="coupon_freq"):
        sp.bootstrap_zero_rates_from_par_yields([1.0], [0.03], coupon_freq=0)


def test_bootstrap_rejects_mismatched_lengths():
    """Ensures par maturities and par yields arrays must match in length."""
    with pytest.raises(ValueError, match="same length"):
        sp.bootstrap_zero_rates_from_par_yields([1.0, 2.0], [0.03])


def test_bootstrap_produces_positive_discount_factors():
    """Smoke check: bootstrapping from a simple Treasury-style par curve yields sensible discount factors (0,1]."""
    mats = np.array([0.5, 1.0, 2.0, 5.0, 10.0])
    par = np.array([0.04, 0.042, 0.045, 0.048, 0.05])
    t_nodes, zero_rates = sp.bootstrap_zero_rates_from_par_yields(mats, par, coupon_freq=2)
    assert len(t_nodes) == len(zero_rates) > 0
    df = np.exp(-zero_rates * t_nodes)
    assert np.all(df > 0) and np.all(df <= 1.0)


# --- MortalityTableQx ---


def test_mortality_qx_load_csv_and_sort(tmp_path):
    """Confirms mortality q_x can be loaded from CSV and ages are sorted for lookup."""
    p = tmp_path / "qx.csv"
    p.write_text("age,qx\n70,0.02\n65,0.01\n", encoding="utf-8")
    mt = sp.MortalityTableQx.load_qx_csv(str(p))
    np.testing.assert_array_equal(mt.ages, np.array([65, 70]))
    np.testing.assert_array_equal(mt.qx, np.array([0.01, 0.02]))


def test_mortality_qx_at_int_age_out_of_range():
    """Ensures looking up q_x outside the table range raises a clear error."""
    ages = np.arange(65, 71)
    qx = np.full_like(ages, 0.01, dtype=float)
    mt = sp.MortalityTableQx(ages, qx)
    with pytest.raises(ValueError, match="outside"):
        mt.qx_at_int_age(64)
    with pytest.raises(ValueError, match="outside"):
        mt.qx_at_int_age(71)


def test_mortality_monthly_survival_monotone(tmp_path):
    """Checks monthly survival probabilities stay between 0 and 1 and do not increase over time."""
    ages = np.arange(65, 110)
    qx = np.clip(0.01 + 0.001 * (ages - 65), 1e-6, 0.35)
    mt = sp.MortalityTableQx(ages, qx)
    S = mt.monthly_survival_to_payment(issue_age=65, n_months=12)
    assert S.shape == (12,)
    assert np.all(S >= 0) and np.all(S <= 1.0)
    assert np.all(np.diff(S) <= 0)


def test_mortality_monthly_survival_rejects_non_positive_n_months():
    """Ensures the monthly survival path rejects a non-positive number of months."""
    mt = sp.MortalityTableQx(np.arange(20, 90), np.full(70, 0.01))
    with pytest.raises(ValueError, match="n_months"):
        mt.monthly_survival_to_payment(issue_age=30, n_months=0)


# --- MortalityTableRP2014MP2016 ---


def _synthetic_mp2016_zero_improvement():
    ages = np.arange(50, 71, dtype=int)
    years = np.arange(2010, 2031, dtype=int)
    i_matrix = np.zeros((len(ages), len(years)), dtype=float)
    base_ages = np.arange(50, 90, dtype=int)
    base_qx = np.full_like(base_ages, 0.02, dtype=float)
    base = sp.MortalityTableQx(base_ages, base_qx)
    return sp.MortalityTableRP2014MP2016(
        base_qx_2014=base,
        mp2016_ages=ages,
        mp2016_years=years,
        mp2016_i_matrix=i_matrix,
        base_year=2014,
    )


def test_rp2014_mp2016_zero_improvement_matches_base_qx():
    """With all MP improvement rates set to zero, projected q_x should match the base RP-2014 table."""
    m = _synthetic_mp2016_zero_improvement()
    assert m.qx_at_int_age_and_calendar_year(age_int=65, calendar_year=2014) == pytest.approx(0.02)
    assert m.qx_at_int_age_and_calendar_year(age_int=65, calendar_year=2025) == pytest.approx(0.02)


def test_rp2014_mp2016_monthly_survival_shape_and_bounds():
    """Checks RP+MP monthly survival produces the right length and probabilities in [0,1] for a short horizon."""
    m = _synthetic_mp2016_zero_improvement()
    S = m.monthly_survival_to_payment(issue_age=65, n_months=24, valuation_year=2024)
    assert S.shape == (24,)
    assert np.all(S >= 0) and np.all(S <= 1.0)


# --- ExpenseAssumptions ---


def test_expense_assumptions_load_csv(tmp_path):
    """Verifies expense assumptions load from a key/value CSV, including percent-to-decimal for premium load."""
    p = tmp_path / "exp.csv"
    rows = [
        {"key": "policy_expense_dollars", "value": 50.0, "unit": "usd"},
        {"key": "premium_expense_rate", "value": 2.0, "unit": "percent"},
        {"key": "monthly_expense_dollars", "value": 3.0, "unit": "usd"},
    ]
    pd.DataFrame(rows).to_csv(p, index=False)
    ex = sp.ExpenseAssumptions.load_from_csv(str(p))
    assert ex.policy_expense_dollars == 50.0
    assert ex.premium_expense_rate == pytest.approx(0.02)
    assert ex.monthly_expense_dollars == 3.0


# --- price_spia_single_premium ---


def _minimal_mortality():
    ages = np.arange(0, 121, dtype=int)
    qx = np.clip(0.005 + ages * 1e-5, 1e-6, 0.4)
    return sp.MortalityTableQx(ages, qx)


def test_price_spia_annuity_factor_and_zero_expenses():
    """Integration: with no expenses, single premium equals PV of benefits and matches survival×discount sum."""
    contract = sp.SPIAContract(issue_age=65, sex="male", benefit_annual=120_000.0)
    yc = sp.YieldCurve.from_flat_rate(0.04)
    mort = _minimal_mortality()
    zero_ex = sp.ExpenseAssumptions(0.0, 0.0, 0.0)
    res = sp.price_spia_single_premium(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=80,
        expenses=zero_ex,
    )
    af = float(np.sum(res.survival_to_payment * res.discount_factors))
    assert res.annuity_factor == pytest.approx(af, rel=1e-9)
    assert res.single_premium == pytest.approx(res.pv_benefit, rel=1e-9)


def test_price_spia_premium_load_formula():
    """Integration: priced premium solves premium = (policy expense + PV benefits) / (1 − premium load rate)."""
    contract = sp.SPIAContract(issue_age=65, sex="male", benefit_annual=60_000.0)
    yc = sp.YieldCurve.from_flat_rate(0.03)
    mort = _minimal_mortality()
    ex = sp.ExpenseAssumptions(policy_expense_dollars=200.0, premium_expense_rate=0.1, monthly_expense_dollars=0.0)
    res = sp.price_spia_single_premium(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=90,
        expenses=ex,
    )
    b_month = contract.benefit_annual / 12.0
    pv_total = float(b_month * res.annuity_factor)
    expected_prem = (200.0 + pv_total) / (1.0 - 0.1)
    assert res.single_premium == pytest.approx(expected_prem, rel=1e-6)


def test_price_spia_rejects_non_monthly_payment_freq():
    """Ensures pricing rejects non-monthly payment frequencies (scaffold only supports 12 per year)."""
    contract = sp.SPIAContract(
        issue_age=65,
        sex="male",
        benefit_annual=100_000.0,
        payment_freq_per_year=4,
    )
    with pytest.raises(ValueError, match="monthly"):
        sp.price_spia_single_premium(
            contract=contract,
            yield_curve=sp.YieldCurve.from_flat_rate(0.04),
            mortality=_minimal_mortality(),
            expenses=sp.ExpenseAssumptions(0, 0, 0),
        )


def test_price_spia_rp_mp_requires_valuation_year():
    """Ensures RP-2014 + MP-2016 mortality requires a valuation year (calendar logic)."""
    m = _synthetic_mp2016_zero_improvement()
    contract = sp.SPIAContract(issue_age=65, sex="male", benefit_annual=50_000.0)
    with pytest.raises(ValueError, match="valuation_year"):
        sp.price_spia_single_premium(
            contract=contract,
            yield_curve=sp.YieldCurve.from_flat_rate(0.04),
            mortality=m,
            valuation_year=None,
            expenses=sp.ExpenseAssumptions(0, 0, 0),
        )


def test_price_spia_rejects_premium_expense_rate_ge_one():
    """Ensures a premium expense rate at or above 100% is rejected (would make premium undefined)."""
    ex = sp.ExpenseAssumptions(0.0, 1.0, 0.0)
    contract = sp.SPIAContract(issue_age=65, sex="male", benefit_annual=40_000.0)
    with pytest.raises(ValueError, match="premium_expense_rate"):
        sp.price_spia_single_premium(
            contract=contract,
            yield_curve=sp.YieldCurve.from_flat_rate(0.04),
            mortality=_minimal_mortality(),
            expenses=ex,
        )


# --- Optional: SOA workbooks (local only) ---


@pytest.mark.skipif(
    not Path(sp.DEFAULT_RP2014_XLSX).is_file(),
    reason="SOA RP-2014 workbook not present",
)
def test_load_rp2014_male_healthy_annuitant_qx_smoke():
    """Optional: if the SOA RP-2014 workbook is present, smoke-test extraction of Healthy Annuitant male q_x."""
    mt = sp.load_rp2014_male_healthy_annuitant_qx_2014(sp.DEFAULT_RP2014_XLSX)
    assert len(mt.ages) > 50
    # Terminal ages in SOA tables may use qx = 1.
    assert np.all(mt.qx >= 0) and np.all(mt.qx <= 1)


@pytest.mark.skipif(
    not Path(sp.DEFAULT_MP2016_XLSX).is_file(),
    reason="SOA MP-2016 workbook not present",
)
def test_load_mp2016_male_improvement_smoke():
    """Optional: if the SOA MP-2016 workbook is present, smoke-test loading the male improvement grid."""
    ages, years, mat = sp.load_mp2016_male_improvement_rates_multiplicative(sp.DEFAULT_MP2016_XLSX)
    assert len(ages) > 10 and len(years) > 10
    assert mat.shape == (len(ages), len(years))


@pytest.mark.skipif(
    not Path(sp.DEFAULT_RP2014_MALE_HEALTHY_QX_CSV).is_file()
    or not Path(sp.DEFAULT_MP2016_MALE_IMPROVEMENT_CSV).is_file(),
    reason="Cached RP-2014 / MP-2016 CSV extracts not present",
)
def test_mp2016_period_qx_stops_at_last_published_year():
    """Excel SUMIFS sums MP rows only through the last calendar year on the grid (no terminal repeat)."""
    sp.ensure_rp2014_male_healthy_annuitant_qx_csv(
        rp2014_xlsx_path=sp.DEFAULT_RP2014_XLSX,
        out_csv_path=sp.DEFAULT_RP2014_MALE_HEALTHY_QX_CSV,
    )
    base = sp.MortalityTableQx.load_qx_csv(sp.DEFAULT_RP2014_MALE_HEALTHY_QX_CSV)
    ages, years, mat = sp.ensure_mp2016_male_improvement_csv(
        mp2016_xlsx_path=sp.DEFAULT_MP2016_XLSX,
        out_csv_path=sp.DEFAULT_MP2016_MALE_IMPROVEMENT_CSV,
    )
    mort = sp.MortalityTableRP2014MP2016(
        base_qx_2014=base,
        mp2016_ages=ages,
        mp2016_years=years,
        mp2016_i_matrix=mat,
    )
    y_last = int(years[-1])
    q_near = mort.qx_at_int_age_and_calendar_year(age_int=50, calendar_year=y_last + 2)
    q_far = mort.qx_at_int_age_and_calendar_year(age_int=50, calendar_year=y_last + 40)
    assert q_near == pytest.approx(q_far, rel=1e-12)


# --- index scenario & expense inflation ---


def test_load_index_scenario_rejects_non_contiguous_months(tmp_path):
    """Index CSV months from 0 must be contiguous (no gaps in the prefix)."""
    p = tmp_path / "idx.csv"
    p.write_text("month,sp500_level\n0,100\n2,102\n", encoding="utf-8")
    with pytest.raises(ValueError, match="contiguous"):
        sp.load_index_scenario_monthly_csv(str(p), n_months=3)


def test_load_index_scenario_extends_short_file_with_flat_tail(tmp_path):
    """Shorter contiguous CSV is extended by holding the last level flat (zero forward returns)."""
    p = tmp_path / "idx.csv"
    p.write_text("month,sp500_level\n0,100\n1,110\n", encoding="utf-8")
    s0, lv = sp.load_index_scenario_monthly_csv(str(p), n_months=3)
    assert s0 == 100.0
    assert lv.shape == (3,)
    assert lv[0] == pytest.approx(110.0)
    assert lv[1] == pytest.approx(110.0)
    assert lv[2] == pytest.approx(110.0)


def test_flat_index_zero_inflation_matches_level_benefits():
    """Flat index and zero expense inflation reproduce level-benefit PV vs legacy-style schedule."""
    contract = sp.SPIAContract(issue_age=65, sex="male", benefit_annual=120_000.0)
    yc = sp.YieldCurve.from_flat_rate(0.04)
    mort = _minimal_mortality()
    ex = sp.ExpenseAssumptions(0.0, 0.0, 0.0)
    res = sp.price_spia_single_premium(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=80,
        expenses=ex,
        index_scenario_csv_path=None,
        expense_annual_inflation=0.0,
    )
    b0 = contract.benefit_annual / 12.0
    af = float(np.sum(res.survival_to_payment * res.discount_factors))
    assert res.pv_benefit == pytest.approx(b0 * af, rel=1e-9)
    np.testing.assert_allclose(res.benefit_nominal_scheduled, np.full_like(res.benefit_nominal_scheduled, b0))
    assert np.max(np.abs(res.index_simple_return)) < 1e-12


def test_expense_inflation_increases_pv_expense(tmp_path):
    """Positive expense inflation strictly increases PV of expenses when survival/curve fixed."""
    p = tmp_path / "idx.csv"
    rows = ["month,sp500_level"]
    for m in range(13):
        rows.append(f"{m},100.0")
    p.write_text("\n".join(rows) + "\n", encoding="utf-8")

    contract = sp.SPIAContract(issue_age=40, sex="male", benefit_annual=0.0)
    yc = sp.YieldCurve.from_flat_rate(0.03)
    ages = np.arange(0, 121, dtype=int)
    qx = np.full_like(ages, 0.01, dtype=float)
    mort = sp.MortalityTableQx(ages, qx)
    ex = sp.ExpenseAssumptions(0.0, 0.0, 50.0)

    r0 = sp.price_spia_single_premium(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=41,
        expenses=ex,
        index_scenario_csv_path=str(p),
        expense_annual_inflation=0.0,
    )
    r1 = sp.price_spia_single_premium(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=41,
        expenses=ex,
        index_scenario_csv_path=str(p),
        expense_annual_inflation=0.12,
    )
    assert r1.pv_monthly_expenses > r0.pv_monthly_expenses
    assert r1.expense_nominal_scheduled[-1] > r1.expense_nominal_scheduled[0]


def test_return_indexation_doubles_with_doubling_index(tmp_path):
    """If index doubles every month, scheduled benefit doubles each month after the first."""
    p = tmp_path / "idx.csv"
    rows = ["month,sp500_level"]
    for m in range(13):
        rows.append(f"{m},{100.0 * (2.0 ** m)}")
    p.write_text("\n".join(rows) + "\n", encoding="utf-8")

    contract = sp.SPIAContract(issue_age=30, sex="male", benefit_annual=120_000.0)
    yc = sp.YieldCurve.from_flat_rate(0.02)
    ages = np.arange(0, 121, dtype=int)
    qx = np.full_like(ages, 0.001, dtype=float)
    mort = sp.MortalityTableQx(ages, qx)
    ex = sp.ExpenseAssumptions(0.0, 0.0, 0.0)
    res = sp.price_spia_single_premium(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=31,
        expenses=ex,
        index_scenario_csv_path=str(p),
        expense_annual_inflation=0.0,
    )
    b = res.benefit_nominal_scheduled
    assert b.shape == (12,)
    assert b[0] == pytest.approx(10_000.0 * 200.0 / 100.0)  # base * S1/S0
    for k in range(1, 12):
        assert b[k] == pytest.approx(2.0 * b[k - 1])


def test_yield_curve_twist_short_equals_short_bps_at_zero():
    """Twist add-on at 0y maturity node must equal bps_short."""
    yc = sp.YieldCurve(
        maturities_years=np.array([0.0, 5.0, 20.0], dtype=float),
        zero_rates=np.array([0.03, 0.03, 0.03], dtype=float),
    )
    out = sp.yield_curve_twist_linear_bps(yc, bps_short=10.0, bps_long=0.0, pivot_years=5.0)
    assert out.zero_rates[0] == pytest.approx(0.03 + 10.0 / 10000.0)


def test_run_alm_projection_smoke_matches_horizon():
    """ALM paths align with pricing horizon and respect liability cashflow override length."""
    contract, yc, mort, ex = _mc_contract_and_setup()
    res = sp.price_spia_single_premium(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=75,
        spread=0.0,
        expenses=ex,
        expense_annual_inflation=0.0,
    )
    asm = sp.ALMAssumptions(
        allocation=sp.alm_default_allocation_spec(),
        rebalance_band=0.10,
        rebalance_frequency_months=1,
        reinvest_rule="hold_cash",
        disinvest_rule="shortest_first",
        liquidity_near_liquid_years=0.25,
    )
    alm = sp.run_alm_projection(pricing=res, yield_curve=yc, spread=0.0, assumptions=asm)
    n = res.months.size
    assert alm.asset_market_value.shape == (n,)
    assert alm.bucket_asset_mv.shape == (len(asm.allocation.buckets), n)
    assert np.isfinite(alm.pv01_net)
    assert np.isfinite(alm.duration_gap)

    cf2 = np.asarray(res.expected_total_cashflows, dtype=float) * 1.05
    alm2 = sp.run_alm_projection(
        pricing=res,
        yield_curve=yc,
        spread=0.0,
        assumptions=asm,
        liability_cashflows=cf2,
    )
    assert alm2.liability_pv[0] > alm.liability_pv[0]


def test_alm_pro_rata_refills_matured_slots_cash_near_target():
    """Matured ladder slots must redeploy at nominal tenors; with no outflows cash stays near target weight."""
    contract, yc, mort, ex = _mc_contract_and_setup()
    res = sp.price_spia_single_premium(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=75,
        spread=0.0,
        expenses=ex,
        expense_annual_inflation=0.0,
    )
    cfz = np.zeros_like(res.expected_total_cashflows, dtype=float)
    asm = sp.ALMAssumptions(
        allocation=sp.alm_default_allocation_spec(),
        rebalance_band=0.05,
        rebalance_frequency_months=1,
        reinvest_rule="pro_rata",
        disinvest_rule="shortest_first",
        liquidity_near_liquid_years=0.25,
    )
    alm = sp.run_alm_projection(
        pricing=res,
        yield_curve=yc,
        spread=0.0,
        assumptions=asm,
        liability_cashflows=cfz,
    )
    w0 = float(asm.allocation.weights[0])
    share = alm.bucket_asset_mv[0, :] / alm.asset_market_value
    assert float(np.max(share)) <= w0 + 0.002
    assert float(np.min(share)) >= w0 - 0.002


def test_alm_pro_rata_reinvest_prioritizes_underweights():
    """Excess cash from maturities should buy underweight buckets before overweight ones."""
    yc = sp.YieldCurve.from_flat_rate(0.0)
    alloc = sp.alm_default_allocation_spec()
    w = np.asarray(alloc.weights, dtype=float)

    faces = np.array([600.0, 100.0, 100.0, 100.0, 100.0], dtype=float)
    t_rem = np.array([1.0, 3.0, 5.0, 10.0, 20.0], dtype=float)
    cash = 200.0
    nominal_tenors = np.array([1.0, 3.0, 5.0, 10.0, 20.0], dtype=float)

    cash2, faces2 = sp._alm_micro_reinvest_pro_rata(
        cash=cash,
        faces=faces,
        t_rem=t_rem.copy(),
        w=w,
        yield_curve=yc,
        spread=0.0,
        nominal_tenors=nominal_tenors,
    )

    delta_faces = faces2 - faces
    # Bucket 0 starts overweight vs target and should not receive new buys.
    assert delta_faces[0] <= 1e-9
    # Underweight buckets should receive purchases.
    assert np.all(delta_faces[1:] > 0.0)
    # Reinvestment should keep cash near target cash weight.
    aum2 = float(cash2 + np.sum(faces2))
    assert abs(cash2 - float(w[0] * aum2)) <= 1e-6


def test_liability_pv_cashflows_length_guard():
    """Mismatched cashflows array must raise."""
    contract, yc, mort, ex = _mc_contract_and_setup()
    res = sp.price_spia_single_premium(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=75,
        spread=0.0,
        expenses=ex,
        expense_annual_inflation=0.0,
    )
    bad = np.ones(3, dtype=float)
    with pytest.raises(ValueError, match="cashflows length"):
        sp.liability_pv_after_paid_months(res, yc, 0.0, -1, cashflows=bad)


def test_yield_curve_twist_rejects_empty_curve():
    empty = sp.YieldCurve(maturities_years=np.array([], dtype=float), zero_rates=np.array([], dtype=float))
    with pytest.raises(ValueError, match="non-empty"):
        sp.yield_curve_twist_linear_bps(empty, bps_short=1.0, bps_long=2.0)


def test_alm_assumptions_validates_rebalance_band():
    alloc = sp.alm_default_allocation_spec()
    with pytest.raises(ValueError, match="rebalance_band"):
        sp.ALMAssumptions(
            allocation=alloc,
            rebalance_band=1.5,
            rebalance_frequency_months=1,
            reinvest_rule="hold_cash",
            disinvest_rule="shortest_first",
        )


def test_alm_assumptions_validates_rebalance_policy():
    alloc = sp.alm_default_allocation_spec()
    with pytest.raises(ValueError, match="rebalance_policy"):
        sp.ALMAssumptions(
            allocation=alloc,
            rebalance_band=0.05,
            rebalance_frequency_months=1,
            reinvest_rule="hold_cash",
            disinvest_rule="shortest_first",
            rebalance_policy="bad_policy",  # type: ignore[arg-type]
        )


def test_alm_assumptions_validates_borrowing_inputs():
    alloc = sp.alm_default_allocation_spec()
    with pytest.raises(ValueError, match="borrowing_policy"):
        sp.ALMAssumptions(
            allocation=alloc,
            rebalance_band=0.05,
            rebalance_frequency_months=1,
            reinvest_rule="hold_cash",
            disinvest_rule="shortest_first",
            borrowing_policy="bad",  # type: ignore[arg-type]
        )
    with pytest.raises(ValueError, match="borrowing_rate_annual"):
        sp.ALMAssumptions(
            allocation=alloc,
            rebalance_band=0.05,
            rebalance_frequency_months=1,
            reinvest_rule="hold_cash",
            disinvest_rule="shortest_first",
            borrowing_rate_annual=-0.01,
        )
    with pytest.raises(ValueError, match="borrowing_rate_mode"):
        sp.ALMAssumptions(
            allocation=alloc,
            rebalance_band=0.05,
            rebalance_frequency_months=1,
            reinvest_rule="hold_cash",
            disinvest_rule="shortest_first",
            borrowing_rate_mode="bad_mode",  # type: ignore[arg-type]
        )
    with pytest.raises(ValueError, match="borrowing_spread_annual"):
        sp.ALMAssumptions(
            allocation=alloc,
            rebalance_band=0.05,
            rebalance_frequency_months=1,
            reinvest_rule="hold_cash",
            disinvest_rule="shortest_first",
            borrowing_spread_annual=-0.01,
        )
    with pytest.raises(ValueError, match="borrowing_rate_tenor_years"):
        sp.ALMAssumptions(
            allocation=alloc,
            rebalance_band=0.05,
            rebalance_frequency_months=1,
            reinvest_rule="hold_cash",
            disinvest_rule="shortest_first",
            borrowing_rate_tenor_years=0.0,
        )


def test_alm_borrowing_policy_changes_sales_behavior():
    contract, yc, mort, ex = _mc_contract_and_setup()
    res = sp.price_spia_single_premium(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=75,
        spread=0.0,
        expenses=ex,
        expense_annual_inflation=0.0,
    )
    # Force large month-1 outflow to create immediate cash shortfall.
    cf = np.asarray(res.expected_total_cashflows, dtype=float).copy()
    cf[0] = float(res.single_premium) * 0.30

    common = dict(
        allocation=sp.alm_default_allocation_spec(),
        rebalance_band=0.10,
        rebalance_frequency_months=1,
        reinvest_rule="hold_cash",
        disinvest_rule="shortest_first",
        rebalance_policy="liquidity_only",
        borrowing_rate_mode="fixed",
        borrowing_rate_annual=0.05,
        liquidity_near_liquid_years=0.25,
    )
    asm_borrow_first = sp.ALMAssumptions(
        **common,
        borrowing_policy="borrow_before_selling",
    )
    asm_sell_first = sp.ALMAssumptions(
        **common,
        borrowing_policy="borrow_after_assets_insufficient",
    )
    out_borrow_first = sp.run_alm_projection(
        pricing=res,
        yield_curve=yc,
        spread=0.0,
        assumptions=asm_borrow_first,
        liability_cashflows=cf,
    )
    out_sell_first = sp.run_alm_projection(
        pricing=res,
        yield_curve=yc,
        spread=0.0,
        assumptions=asm_sell_first,
        liability_cashflows=cf,
    )
    # Borrow-first should preserve bond MV better in early month and show positive borrowing balance.
    assert float(np.sum(out_borrow_first.bucket_asset_mv[1:, 0])) >= float(np.sum(out_sell_first.bucket_asset_mv[1:, 0]))
    assert float(out_borrow_first.borrowing_balance[0]) >= 0.0


def test_alm_scenario_linked_borrow_rate_tracks_curve_level():
    contract, yc, mort, ex = _mc_contract_and_setup()
    res = sp.price_spia_single_premium(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=75,
        spread=0.0,
        expenses=ex,
        expense_annual_inflation=0.0,
    )
    cf = np.asarray(res.expected_total_cashflows, dtype=float).copy()
    cf[0] = float(res.single_premium) * 0.30
    alloc = sp.alm_default_allocation_spec()
    asm = sp.ALMAssumptions(
        allocation=alloc,
        rebalance_band=0.10,
        rebalance_frequency_months=1,
        reinvest_rule="hold_cash",
        disinvest_rule="shortest_first",
        rebalance_policy="liquidity_only",
        borrowing_policy="borrow_before_selling",
        borrowing_rate_mode="scenario_linked",
        borrowing_spread_annual=0.0,
        liquidity_near_liquid_years=0.25,
    )
    yc_low = sp.YieldCurve.from_flat_rate(0.01)
    yc_high = sp.YieldCurve.from_flat_rate(0.08)
    out_low = sp.run_alm_projection(
        pricing=res,
        yield_curve=yc_low,
        spread=0.0,
        assumptions=asm,
        liability_cashflows=cf,
    )
    out_high = sp.run_alm_projection(
        pricing=res,
        yield_curve=yc_high,
        spread=0.0,
        assumptions=asm,
        liability_cashflows=cf,
    )
    # With scenario-linked borrowing and same initial debt, higher curve => higher debt accrual by month 2.
    assert float(out_high.borrowing_balance[1]) > float(out_low.borrowing_balance[1])


def test_alm_scenario_linked_borrow_rate_respects_selected_tenor():
    contract, yc, mort, ex = _mc_contract_and_setup()
    res = sp.price_spia_single_premium(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=75,
        spread=0.0,
        expenses=ex,
        expense_annual_inflation=0.0,
    )
    cf = np.asarray(res.expected_total_cashflows, dtype=float).copy()
    cf[0] = float(res.single_premium) * 0.30

    # Upward-sloping curve: longer tenor should imply higher borrowing base rate.
    yc_up = sp.YieldCurve(
        maturities_years=np.array([0.0, 0.25, 1.0, 3.0, 5.0], dtype=float),
        zero_rates=np.array([0.01, 0.02, 0.03, 0.045, 0.05], dtype=float),
    )
    common = dict(
        allocation=sp.alm_default_allocation_spec(),
        rebalance_band=0.10,
        rebalance_frequency_months=1,
        reinvest_rule="hold_cash",
        disinvest_rule="shortest_first",
        rebalance_policy="liquidity_only",
        borrowing_policy="borrow_before_selling",
        borrowing_rate_mode="scenario_linked",
        borrowing_spread_annual=0.0,
        liquidity_near_liquid_years=0.25,
    )
    asm_short = sp.ALMAssumptions(**common, borrowing_rate_tenor_years=0.25)
    asm_long = sp.ALMAssumptions(**common, borrowing_rate_tenor_years=3.0)
    out_short = sp.run_alm_projection(
        pricing=res,
        yield_curve=yc_up,
        spread=0.0,
        assumptions=asm_short,
        liability_cashflows=cf,
    )
    out_long = sp.run_alm_projection(
        pricing=res,
        yield_curve=yc_up,
        spread=0.0,
        assumptions=asm_long,
        liability_cashflows=cf,
    )
    assert float(out_long.borrowing_balance[1]) > float(out_short.borrowing_balance[1])

