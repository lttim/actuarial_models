"""Unit tests for spia_projection (SPIA pricing scaffold)."""

from __future__ import annotations

import math
from pathlib import Path

import numpy as np
import pandas as pd
import pytest

import spia_projection as sp


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
