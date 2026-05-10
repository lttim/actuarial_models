"""Regression: Term what-if must use Term pricing horizon, not SPIA full life horizon."""

from __future__ import annotations

import numpy as np
import pytest

from annuity_model import pricing_projection as sp
from annuity_model import rila_projection as rp
from annuity_model import term_projection as tp
from annuity_model.pricing_ui import (
    build_alm_pricing_for_mc_scenario,
    compute_what_if_term_shocked_pricing,
)
from annuity_model.product_registry import ProductType


def _term_case() -> tuple[tp.TermLifeContract, sp.YieldCurve, sp.MortalityTableQx]:
    contract = tp.TermLifeContract(
        issue_age=45,
        sex="male",
        death_benefit=250_000.0,
        monthly_premium=250.0,
        term_years=20,
    )
    yc = sp.YieldCurve.from_flat_rate(0.04)
    ages = np.arange(0, 121, dtype=int)
    qx = np.full_like(ages, 0.01, dtype=float)
    mort = sp.MortalityTableQx(ages, qx)
    return contract, yc, mort


def test_what_if_term_shocked_pricing_matches_month_count_and_avoids_spia_horizon_trap():
    """
    Historical bug: what-if called SPIA pricing with Term-length index paths, raising
    ValueError('index_levels_payment must have shape (540,)') while Term has 240 months.
    """
    contract, yc, mort = _term_case()
    base = tp.price_term_life_level_monthly(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=110,
        spread=0.0,
        valuation_year=None,
    )
    n_term = int(base.months.size)
    assert n_term == 240

    spia = sp.SPIAContract(
        issue_age=contract.issue_age,
        sex=contract.sex,
        benefit_annual=float(contract.death_benefit),
    )
    idx_wrong_len = np.ones(n_term, dtype=float) * 100.0
    with pytest.raises(ValueError, match="index_levels_payment must have shape"):
        sp.price_spia_single_premium(
            contract=spia,
            yield_curve=yc,
            mortality=mort,
            horizon_age=110,
            spread=0.0,
            valuation_year=None,
            expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
            index_s0=100.0,
            index_levels_payment=idx_wrong_len,
            expense_annual_inflation=0.0,
        )

    shocked = compute_what_if_term_shocked_pricing(
        base_contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=110,
        spread=0.0,
        valuation_year=None,
        term_monthly_premium_mult=1.0,
    )
    assert shocked.months.size == n_term
    np.testing.assert_allclose(shocked.single_premium, base.single_premium, rtol=0.0, atol=1e-9)

    shocked_looser = compute_what_if_term_shocked_pricing(
        base_contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=110,
        spread=0.0,
        valuation_year=None,
        term_monthly_premium_mult=0.8,
    )
    assert shocked_looser.months.size == n_term
    assert float(shocked_looser.single_premium) > float(shocked.single_premium)


def test_build_alm_mc_scenario_skips_spia_repricer_for_term_product():
    contract, yc, mort = _term_case()
    term_res = tp.price_term_life_level_monthly(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=110,
        spread=0.0,
        valuation_year=None,
    )
    expenses = sp.ExpenseAssumptions(0.0, 0.0, 0.0)
    out = build_alm_pricing_for_mc_scenario(
        product_type=ProductType.TERM_LIFE,
        scenario_source="MC simulation (single path)",
        baseline_pricing=term_res,
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=110,
        spread=0.0,
        valuation_year=None,
        expenses=expenses,
        expense_annual_inflation=0.0,
        mc_n_sims=100,
        mc_seed=42,
        mc_scenario_idx=0,
        mc_params={"s0": 100.0, "annual_drift": 0.06, "annual_vol": 0.15},
    )
    assert out is term_res


def test_build_alm_mc_scenario_reprices_rila_on_single_path():
    contract = rp.RILAContract(
        issue_age=55,
        sex="male",
        participation=1.0,
        cap=0.1,
        floor=0.0,
        rider_fee_annual=0.0,
    )
    yc = sp.YieldCurve.from_flat_rate(0.04)
    ages = np.arange(0, 121, dtype=int)
    qx = np.full_like(ages, 0.015, dtype=float)
    mort = sp.MortalityTableQx(ages, qx)
    # RILA single-premium solve requires a positive expense numerator when
    # premium_expense_rate is zero; all-zero expenses yield a collapsed premium.
    expenses = sp.ExpenseAssumptions(1.0, 0.0, 0.0)
    base = rp.price_rila_single_premium(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=75,
        spread=0.0,
        valuation_year=None,
        expenses=expenses,
        index_scenario_csv_path=None,
        expense_annual_inflation=0.0,
    )
    out = build_alm_pricing_for_mc_scenario(
        product_type=ProductType.RILA,
        scenario_source="MC simulation (single path)",
        baseline_pricing=base,
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=75,
        spread=0.0,
        valuation_year=None,
        expenses=expenses,
        expense_annual_inflation=0.0,
        mc_n_sims=50,
        mc_seed=7,
        mc_scenario_idx=3,
        mc_params={"s0": 100.0, "annual_drift": 0.06, "annual_vol": 0.15},
    )
    assert out is not base
    assert isinstance(out, rp.RILAProjectionResult)
    assert out.months.size == base.months.size
