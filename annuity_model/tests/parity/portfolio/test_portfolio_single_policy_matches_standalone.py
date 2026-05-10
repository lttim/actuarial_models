"""One-policy portfolio run matches standalone ``adapter.price`` for the same scenario."""

from __future__ import annotations

import numpy as np
import pandas as pd
import pytest

from annuity_model.inforce_io import load_policy_inputs_from_csv_from_dataframe
from annuity_model.portfolio import Portfolio
from annuity_model.portfolio_runner import run_portfolio
from annuity_model.pricing_run_form_state import (
    build_run_form_seed_defaults,
    default_inforce_scratch_row,
)
from annuity_model.pricing_scenario_materialize import (
    ANN_MODEL_ROOT,
    run_scenario_from_pricing_seeds,
)
from annuity_model.product_registry import ProductType, get_product_adapter

pytestmark = pytest.mark.parity

_IMPLEMENTED = (
    ProductType.SPIA,
    ProductType.TERM_LIFE,
    ProductType.RILA,
    ProductType.MYGA,
    ProductType.FIA,
    ProductType.VARIABLE_ANNUITY,
    ProductType.WHOLE_LIFE,
    ProductType.UNIVERSAL_LIFE,
    ProductType.INDEXED_UL,
    ProductType.VARIABLE_UL,
)


@pytest.mark.parametrize("pt", _IMPLEMENTED)
def test_single_policy_portfolio_matches_standalone_pricing(pt: ProductType) -> None:
    row = default_inforce_scratch_row(pt)
    row["policy_id"] = f"single-{pt.value}"
    policies = load_policy_inputs_from_csv_from_dataframe(pd.DataFrame([row]))
    assert len(policies) == 1
    pol = policies[0]

    seeds = build_run_form_seed_defaults(
        product_default=pt.value,
        saved_inputs={},
        meta={},
        default_product_type=pt,
    )
    scen = run_scenario_from_pricing_seeds(
        seeds, default_product_type=pt, sex="male", repo_root=ANN_MODEL_ROOT
    )

    adapter = get_product_adapter(pt)
    standalone = adapter.price(
        contract=pol.contract,
        yield_curve=scen.yield_curve,
        mortality=scen.mortality,
        horizon_age=scen.horizon_age,
        spread=scen.spread,
        valuation_year=scen.valuation_year,
        expenses=scen.expenses,
        expenses_csv_path=scen.expenses_csv_path,
        index_scenario_csv_path=scen.index_scenario_csv_path,
        expense_annual_inflation=scen.expense_annual_inflation,
    )

    res = run_portfolio(portfolio=Portfolio(policies=policies), scenario=scen)
    assert len(res.policy_results) == 1
    pr = res.policy_results[0]

    sp0 = float(getattr(standalone, "single_premium"))
    sp1 = float(getattr(pr.pricing, "single_premium"))
    np.testing.assert_allclose(sp1, sp0, rtol=0, atol=1e-6)

    pv0 = float(getattr(standalone, "pv_benefit"))
    pv1 = float(getattr(pr.pricing, "pv_benefit"))
    np.testing.assert_allclose(pv1, pv0, rtol=0, atol=1e-4)

    cf0 = np.asarray(standalone.expected_total_cashflows, dtype=float)
    cf1 = np.asarray(pr.pricing.expected_total_cashflows, dtype=float)
    np.testing.assert_allclose(cf1, cf0, rtol=0, atol=1e-9)
