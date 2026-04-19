"""Regression matrix: every (product, scenario, surface) cell must be green or skipped.

Why this test exists
--------------------

The platform exposes ~7 distinct *surfaces* per product (deterministic
pricing, Monte Carlo pricing, liability-path conversion, ALM projection,
Excel workbook build, Excel workbook static validation, UI metric
formatter). For every product (SPIA / Term / RILA / future) and every
"flavor" of input (baseline / minimum-horizon / etc.), each of those
surfaces must either:

  * have a smoke test that actually exercises it end-to-end, OR
  * be explicitly skipped with a documented reason in
    ``EXPECTED_SKIPS`` below.

A *missing* cell -- a (product, scenario, surface) tuple that nobody
ever exercises and nobody explicitly skipped -- is exactly the failure
mode that lets a new product ship with broken Monte Carlo or a
silently-degraded Excel build.

This test is deliberately the *only* place where the (product × scenario
× surface) cube is enumerated. If you add a new product, you do NOT need
to write more parametrize entries: ``PRODUCT_FIXTURES`` derives products
from the live ``implemented_product_types()`` list, so the matrix
auto-grows. You only have to add the fixture-builder for the new product
in ``_build_<product>_fixture()`` and (if applicable) opt-out of any
surface in ``EXPECTED_SKIPS``.

What this test is NOT
---------------------

* This is *not* a parity test (``tests/parity/`` owns those).
* This is *not* a Monte Carlo statistical test (``tests/test_pricing_projection.py``
  owns those).
* This is *not* an Excel-validator unit test (``tests/test_excel_export_validation.py``).

It is a *coverage matrix*: a single, exhaustive enumeration that
guarantees no surface silently goes untested for any product.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Callable

import numpy as np
import pytest

import build_fia_excel_workbook as build_fia_xl
import build_iul_excel_workbook as build_iul_xl
import build_myga_excel_workbook as build_myga_xl
import build_pricing_excel_workbook as build_spia_xl
import build_rila_excel_workbook as build_rila_xl
import build_term_excel_workbook as build_term_xl
import build_ul_excel_workbook as build_ul_xl
import build_va_excel_workbook as build_va_xl
import build_vul_excel_workbook as build_vul_xl
import build_wl_excel_workbook as build_wl_xl
import excel_workbook_validator
import fia_projection as fp
import iul_projection as iul_proj
import myga_projection as my
import pricing_projection as sp
import rila_projection as rp
import term_projection as tp
import ul_projection as ul_proj
import va_projection as va
import vul_projection as vul_proj
import wl_projection as wl
from liability_dispatch import liability_path_for
from product_registry import (
    ProductType,
    _PRICING_METRIC_FORMATTERS,
    implemented_product_types,
    validate_run_inputs,
)


# ---------------------------------------------------------------------------
# Fixture builders -- one per product. Keep these tiny: a fast,
# deterministic, "happy-path" set of inputs. Variation belongs in SCENARIOS,
# not here.
# ---------------------------------------------------------------------------


@dataclass
class ProductFixture:
    """Bag of canonical inputs + cached pricing result for one product."""

    product_type: ProductType
    contract: Any
    yield_curve: sp.YieldCurve
    mortality: sp.MortalityTableQx
    horizon_age: int
    expenses: sp.ExpenseAssumptions
    pricing_kwargs: dict[str, Any]
    # Pre-computed result so the test functions don't have to repeat work
    # across surface checks. Built lazily by `pricing_result()`.
    _result: Any = None


def _flat_yc(rate: float = 0.04) -> sp.YieldCurve:
    return sp.YieldCurve.from_flat_rate(rate)


def _synthetic_mortality(start: float = 0.005, slope: float = 1e-5) -> sp.MortalityTableQx:
    ages = np.arange(0, 121, dtype=int)
    qx = np.clip(start + ages * slope, 1e-6, 0.4)
    return sp.MortalityTableQx(ages, qx)


def _build_spia_fixture(scenario: str) -> ProductFixture:
    horizon_age = 80 if scenario == "baseline" else 70
    contract = sp.SPIAContract(issue_age=65, sex="male", benefit_annual=100_000.0)
    return ProductFixture(
        product_type=ProductType.SPIA,
        contract=contract,
        yield_curve=_flat_yc(),
        mortality=_synthetic_mortality(),
        horizon_age=horizon_age,
        expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
        pricing_kwargs=dict(spread=0.0, valuation_year=None, expense_annual_inflation=0.0),
    )


def _build_term_fixture(scenario: str) -> ProductFixture:
    horizon_age = 60 if scenario == "baseline" else 45
    contract = tp.TermLifeContract(
        issue_age=40,
        sex="male",
        death_benefit=250_000.0,
        monthly_premium=200.0,
        term_years=20,
        premium_mode="level_monthly",
        benefit_timing="eoy_death",
    )
    return ProductFixture(
        product_type=ProductType.TERM_LIFE,
        contract=contract,
        yield_curve=_flat_yc(),
        mortality=_synthetic_mortality(),
        horizon_age=horizon_age,
        expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
        pricing_kwargs=dict(spread=0.0),
    )


def _build_rila_fixture(scenario: str) -> ProductFixture:
    horizon_age = 85 if scenario == "baseline" else 65
    contract = rp.RILAContract(
        issue_age=55,
        sex="male",
        participation=0.85,
        cap=0.09,
        floor=-0.02,
        rider_fee_annual=0.008,
    )
    return ProductFixture(
        product_type=ProductType.RILA,
        contract=contract,
        yield_curve=_flat_yc(0.035),
        mortality=_synthetic_mortality(0.008, 2e-5),
        horizon_age=horizon_age,
        expenses=sp.ExpenseAssumptions(0.0, 0.0, 25.0),
        pricing_kwargs=dict(spread=0.0, valuation_year=None, expense_annual_inflation=0.01),
    )


def _build_myga_fixture(scenario: str) -> ProductFixture:
    horizon_age = 70 if scenario == "baseline" else 65
    contract = my.MYGAContract(
        issue_age=60, sex="male", single_premium=100_000.0,
        declared_rate_annual=0.045, guarantee_years=5,
    )
    return ProductFixture(
        product_type=ProductType.MYGA, contract=contract,
        yield_curve=_flat_yc(0.045), mortality=_synthetic_mortality(),
        horizon_age=horizon_age, expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
        pricing_kwargs=dict(spread=0.0, valuation_year=None, expense_annual_inflation=0.0),
    )


def _build_fia_fixture(scenario: str) -> ProductFixture:
    horizon_age = 70 if scenario == "baseline" else 65
    contract = fp.FIAContract(
        issue_age=60, sex="male", single_premium=100_000.0,
        participation=0.8, cap=0.07, floor=0.0, horizon_years=10,
    )
    return ProductFixture(
        product_type=ProductType.FIA, contract=contract,
        yield_curve=_flat_yc(), mortality=_synthetic_mortality(),
        horizon_age=horizon_age, expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
        pricing_kwargs=dict(spread=0.0, valuation_year=None, expense_annual_inflation=0.0),
    )


def _build_va_fixture(scenario: str) -> ProductFixture:
    horizon_age = 75 if scenario == "baseline" else 65
    contract = va.VAContract(
        issue_age=55, sex="male", single_premium=100_000.0,
        me_charge_annual=0.014, horizon_years=20,
    )
    return ProductFixture(
        product_type=ProductType.VARIABLE_ANNUITY, contract=contract,
        yield_curve=_flat_yc(), mortality=_synthetic_mortality(),
        horizon_age=horizon_age, expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
        pricing_kwargs=dict(spread=0.0, valuation_year=None, expense_annual_inflation=0.0),
    )


def _build_wl_fixture(scenario: str) -> ProductFixture:
    horizon_age = 120 if scenario == "baseline" else 80
    contract = wl.WLContract(
        issue_age=45, sex="male", smoker_class="nonsmoker", face_amount=250_000.0,
    )
    return ProductFixture(
        product_type=ProductType.WHOLE_LIFE, contract=contract,
        yield_curve=_flat_yc(), mortality=_synthetic_mortality(),
        horizon_age=horizon_age, expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
        pricing_kwargs=dict(spread=0.0, valuation_year=None, expense_annual_inflation=0.0),
    )


def _build_ul_fixture(scenario: str) -> ProductFixture:
    horizon_age = 120 if scenario == "baseline" else 80
    contract = ul_proj.ULContract(
        issue_age=45, sex="male", face_amount=250_000.0, single_premium=25_000.0,
    )
    return ProductFixture(
        product_type=ProductType.UNIVERSAL_LIFE, contract=contract,
        yield_curve=_flat_yc(), mortality=_synthetic_mortality(),
        horizon_age=horizon_age, expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
        pricing_kwargs=dict(spread=0.0, valuation_year=None, expense_annual_inflation=0.0),
    )


def _build_iul_fixture(scenario: str) -> ProductFixture:
    horizon_age = 120 if scenario == "baseline" else 80
    contract = iul_proj.IULContract(
        issue_age=45, sex="male", face_amount=250_000.0, single_premium=25_000.0,
        participation=1.0, cap=0.10, floor=0.0,
    )
    return ProductFixture(
        product_type=ProductType.INDEXED_UL, contract=contract,
        yield_curve=_flat_yc(), mortality=_synthetic_mortality(),
        horizon_age=horizon_age, expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
        pricing_kwargs=dict(spread=0.0, valuation_year=None, expense_annual_inflation=0.0),
    )


def _build_vul_fixture(scenario: str) -> ProductFixture:
    horizon_age = 120 if scenario == "baseline" else 80
    contract = vul_proj.VULContract(
        issue_age=45, sex="male", face_amount=250_000.0, single_premium=25_000.0,
    )
    return ProductFixture(
        product_type=ProductType.VARIABLE_UL, contract=contract,
        yield_curve=_flat_yc(), mortality=_synthetic_mortality(),
        horizon_age=horizon_age, expenses=sp.ExpenseAssumptions(0.0, 0.0, 0.0),
        pricing_kwargs=dict(spread=0.0, valuation_year=None, expense_annual_inflation=0.0),
    )


_FIXTURE_BUILDERS: dict[ProductType, Callable[[str], ProductFixture]] = {
    ProductType.SPIA: _build_spia_fixture,
    ProductType.TERM_LIFE: _build_term_fixture,
    ProductType.RILA: _build_rila_fixture,
    ProductType.MYGA: _build_myga_fixture,
    ProductType.FIA: _build_fia_fixture,
    ProductType.VARIABLE_ANNUITY: _build_va_fixture,
    ProductType.WHOLE_LIFE: _build_wl_fixture,
    ProductType.UNIVERSAL_LIFE: _build_ul_fixture,
    ProductType.INDEXED_UL: _build_iul_fixture,
    ProductType.VARIABLE_UL: _build_vul_fixture,
}


# ---------------------------------------------------------------------------
# Surface runners -- one per surface. Each takes a fixture and either
# returns silently (cell green) or raises (cell red). They MUST NOT be
# allowed to be quietly no-ops; the parametrize list below also enforces
# that every surface name is wired here.
# ---------------------------------------------------------------------------


def _index_levels_for(fix: ProductFixture) -> tuple[float, np.ndarray]:
    """Build a deterministic index path sized to the fixture's effective horizon.

    For products with their own ``horizon_years`` field (FIA / VA / IUL /
    VUL), the engine takes ``min(horizon_age, contract.horizon_years*12)``
    months. We size the array to the maximum of the two so the index
    path is always long enough.
    """
    base_n = int(round((fix.horizon_age - fix.contract.issue_age) * 12))
    contract_n = int(getattr(fix.contract, "horizon_years", 0)) * 12
    n_months = max(1, max(base_n, contract_n))
    rng = np.random.default_rng(42)
    levels = 100.0 * np.cumprod(1.0 + rng.normal(0.004, 0.02, size=n_months))
    return 100.0, levels


def _truncate_levels_for_engine(
    fix: ProductFixture, levels: np.ndarray
) -> np.ndarray:
    """Truncate the index path to the n_months the engine expects.

    Different engines compute n_months differently. To keep _index_levels_for
    generic, we provide the maximum needed length and trim to the engine's
    actual demand here.
    """
    base_n = int(round((fix.horizon_age - fix.contract.issue_age) * 12))
    contract_n = int(getattr(fix.contract, "horizon_years", 0)) * 12
    if fix.product_type is ProductType.RILA:
        # RILA uses (horizon_age - issue_age)*12 only.
        target = base_n
    elif fix.product_type in (
        ProductType.FIA, ProductType.VARIABLE_ANNUITY,
        ProductType.INDEXED_UL, ProductType.VARIABLE_UL,
    ):
        # All four use min(base_n, contract_n).
        target = min(base_n, contract_n) if contract_n > 0 else base_n
    else:
        target = base_n
    return levels[: max(1, target)]


def _price_deterministic(fix: ProductFixture) -> Any:
    """Run the deterministic price() and cache the result on the fixture."""
    if fix.product_type is ProductType.SPIA:
        result = sp.price_spia_single_premium(
            contract=fix.contract, yield_curve=fix.yield_curve,
            mortality=fix.mortality, horizon_age=fix.horizon_age,
            expenses=fix.expenses, **fix.pricing_kwargs,
        )
    elif fix.product_type is ProductType.TERM_LIFE:
        result = tp.price_term_life_level_monthly(
            contract=fix.contract, yield_curve=fix.yield_curve,
            mortality=fix.mortality, horizon_age=fix.horizon_age,
            **fix.pricing_kwargs,
        )
    elif fix.product_type is ProductType.RILA:
        s0, levels = _index_levels_for(fix)
        levels = _truncate_levels_for_engine(fix, levels)
        result = rp.price_rila_single_premium(
            contract=fix.contract, yield_curve=fix.yield_curve,
            mortality=fix.mortality, horizon_age=fix.horizon_age,
            expenses=fix.expenses, index_s0=s0, index_levels_payment=levels,
            **fix.pricing_kwargs,
        )
    elif fix.product_type is ProductType.MYGA:
        result = my.price_myga_single_premium(
            contract=fix.contract, yield_curve=fix.yield_curve,
            mortality=fix.mortality, horizon_age=fix.horizon_age,
            expenses=fix.expenses, **fix.pricing_kwargs,
        )
    elif fix.product_type is ProductType.FIA:
        s0, levels = _index_levels_for(fix)
        levels = _truncate_levels_for_engine(fix, levels)
        result = fp.price_fia_single_premium(
            contract=fix.contract, yield_curve=fix.yield_curve,
            mortality=fix.mortality, horizon_age=fix.horizon_age,
            expenses=fix.expenses, index_s0=s0, index_levels_payment=levels,
            **fix.pricing_kwargs,
        )
    elif fix.product_type is ProductType.VARIABLE_ANNUITY:
        s0, levels = _index_levels_for(fix)
        levels = _truncate_levels_for_engine(fix, levels)
        result = va.price_va_single_premium(
            contract=fix.contract, yield_curve=fix.yield_curve,
            mortality=fix.mortality, horizon_age=fix.horizon_age,
            expenses=fix.expenses, index_s0=s0, index_levels_payment=levels,
            **fix.pricing_kwargs,
        )
    elif fix.product_type is ProductType.WHOLE_LIFE:
        result = wl.price_wl_single_premium(
            contract=fix.contract, yield_curve=fix.yield_curve,
            mortality=fix.mortality, horizon_age=fix.horizon_age,
            expenses=fix.expenses, **fix.pricing_kwargs,
        )
    elif fix.product_type is ProductType.UNIVERSAL_LIFE:
        result = ul_proj.price_ul_single_premium(
            contract=fix.contract, yield_curve=fix.yield_curve,
            mortality=fix.mortality, horizon_age=fix.horizon_age,
            expenses=fix.expenses, **fix.pricing_kwargs,
        )
    elif fix.product_type is ProductType.INDEXED_UL:
        s0, levels = _index_levels_for(fix)
        levels = _truncate_levels_for_engine(fix, levels)
        result = iul_proj.price_iul_single_premium(
            contract=fix.contract, yield_curve=fix.yield_curve,
            mortality=fix.mortality, horizon_age=fix.horizon_age,
            expenses=fix.expenses, index_s0=s0, index_levels_payment=levels,
            **fix.pricing_kwargs,
        )
    elif fix.product_type is ProductType.VARIABLE_UL:
        s0, levels = _index_levels_for(fix)
        levels = _truncate_levels_for_engine(fix, levels)
        result = vul_proj.price_vul_single_premium(
            contract=fix.contract, yield_curve=fix.yield_curve,
            mortality=fix.mortality, horizon_age=fix.horizon_age,
            expenses=fix.expenses, index_s0=s0, index_levels_payment=levels,
            **fix.pricing_kwargs,
        )
    else:
        raise AssertionError(f"unhandled product {fix.product_type!r}")
    fix._result = result
    return result


def _price_monte_carlo(fix: ProductFixture) -> Any:
    if fix.product_type is ProductType.SPIA:
        return sp.price_spia_single_premium_monte_carlo(
            contract=fix.contract, yield_curve=fix.yield_curve,
            mortality=fix.mortality, horizon_age=fix.horizon_age,
            expenses=fix.expenses, n_sims=8, annual_drift=0.04, annual_vol=0.15,
            seed=0, s0=100.0, **fix.pricing_kwargs,
        )
    if fix.product_type is ProductType.RILA:
        return rp.price_rila_single_premium_monte_carlo(
            contract=fix.contract, yield_curve=fix.yield_curve,
            mortality=fix.mortality, horizon_age=fix.horizon_age,
            expenses=fix.expenses, n_sims=8, annual_drift=0.04, annual_vol=0.15,
            seed=0, s0=100.0, **fix.pricing_kwargs,
        )
    if fix.product_type is ProductType.FIA:
        return fp.price_fia_single_premium_monte_carlo(
            contract=fix.contract, yield_curve=fix.yield_curve,
            mortality=fix.mortality, horizon_age=fix.horizon_age,
            expenses=fix.expenses, n_sims=4, annual_drift=0.04, annual_vol=0.15,
            seed=0, s0=100.0, **fix.pricing_kwargs,
        )
    if fix.product_type is ProductType.VARIABLE_ANNUITY:
        return va.price_va_single_premium_monte_carlo(
            contract=fix.contract, yield_curve=fix.yield_curve,
            mortality=fix.mortality, horizon_age=fix.horizon_age,
            expenses=fix.expenses, n_sims=4, annual_drift=0.04, annual_vol=0.15,
            seed=0, s0=100.0, **fix.pricing_kwargs,
        )
    if fix.product_type is ProductType.INDEXED_UL:
        return iul_proj.price_iul_single_premium_monte_carlo(
            contract=fix.contract, yield_curve=fix.yield_curve,
            mortality=fix.mortality, horizon_age=fix.horizon_age,
            expenses=fix.expenses, n_sims=4, annual_drift=0.04, annual_vol=0.15,
            seed=0, s0=100.0, **fix.pricing_kwargs,
        )
    if fix.product_type is ProductType.VARIABLE_UL:
        return vul_proj.price_vul_single_premium_monte_carlo(
            contract=fix.contract, yield_curve=fix.yield_curve,
            mortality=fix.mortality, horizon_age=fix.horizon_age,
            expenses=fix.expenses, n_sims=4, annual_drift=0.04, annual_vol=0.15,
            seed=0, s0=100.0, **fix.pricing_kwargs,
        )
    raise NotImplementedError(
        f"Monte Carlo not exercised by this matrix for {fix.product_type!r}; "
        "if a MC entry-point exists, add it to _price_monte_carlo() instead "
        "of skipping the cell."
    )


def _liability_path(fix: ProductFixture) -> Any:
    if fix._result is None:
        _price_deterministic(fix)
    return liability_path_for(fix._result)


def _alm_projection(fix: ProductFixture) -> Any:
    """Run a small ALM projection from the cached liability path.

    We deliberately use the liability-path entry point (not the higher-level
    wrapper) so the test exercises exactly the dispatch path the platform
    uses in production.
    """
    if fix._result is None:
        _price_deterministic(fix)
    path = liability_path_for(fix._result)
    asm = sp.ALMAssumptions(
        allocation=sp.alm_default_allocation_spec(),
        rebalance_band=0.10,
        rebalance_frequency_months=1,
        reinvest_rule="hold_cash",
        disinvest_rule="shortest_first",
        rebalance_policy="liquidity_only",
        liquidity_near_liquid_years=0.25,
    )
    return sp.run_alm_projection_from_liability_path(
        liability_path=path,
        yield_curve=fix.yield_curve,
        spread=0.0,
        assumptions=asm,
        initial_asset_market_value=500_000.0,
    )


def _build_excel_workbook(fix: ProductFixture) -> bytes:
    """Build the per-product workbook end-to-end and return raw bytes."""
    if fix._result is None:
        _price_deterministic(fix)
    res = fix._result
    if fix.product_type is ProductType.SPIA:
        spec = build_spia_xl.excel_spec_from_launcher(
            contract=fix.contract,
            yield_curve=fix.yield_curve,
            mortality=fix.mortality,
            horizon_age=fix.horizon_age,
            spread=0.0,
            valuation_year=2025,
            expenses=fix.expenses,
            yield_mode_label="flat",
            mortality_mode_label="synthetic",
            expense_mode_label="manual",
            index_s0=float(res.index_s0),
            index_levels_at_payment=res.index_level_at_payment,
            expense_annual_inflation=0.0,
        )
        snap = build_spia_xl.ExcelPythonSnapshot(
            pv_benefit=float(res.pv_benefit),
            pv_monthly_expenses=float(res.pv_monthly_expenses),
            pv_monthly_total=float(res.pv_benefit + res.pv_monthly_expenses),
            single_premium=float(res.single_premium),
            annuity_factor=float(res.annuity_factor),
        )
        return build_spia_xl.build_workbook_from_spec(spec, out_path=None, python_snapshot=snap)
    if fix.product_type is ProductType.TERM_LIFE:
        spec = build_term_xl.term_excel_spec_from_launcher(
            contract=fix.contract,
            yield_curve=fix.yield_curve,
            mortality=fix.mortality,
            horizon_age=fix.horizon_age,
            spread=0.0,
            valuation_year=2025,
            expenses=fix.expenses,
            yield_mode_label="flat",
            mortality_mode_label="synthetic",
            expense_mode_label="manual",
        )
        return build_term_xl.build_term_workbook_from_spec(spec)
    if fix.product_type is ProductType.RILA:
        spec = build_rila_xl.rila_excel_spec_from_launcher(
            contract=fix.contract,
            yield_curve=fix.yield_curve,
            mortality=fix.mortality,
            horizon_age=fix.horizon_age,
            spread=0.0,
            valuation_year=2025,
            expenses=fix.expenses,
            yield_mode_label="flat",
            mortality_mode_label="synthetic",
            expense_mode_label="manual",
            index_s0=100.0,
            index_levels_at_payment=res.index_level_at_payment,
            expense_annual_inflation=0.01,
        )
        return build_rila_xl.build_rila_workbook_from_spec(spec)
    # Seven new products use a uniform spec/builder shape. Pull the
    # builder + spec_fn from product_excel via _BUILDER_REGISTRY.
    if fix.product_type is ProductType.MYGA:
        spec = build_myga_xl.myga_excel_spec_from_launcher(
            contract=fix.contract, yield_curve=fix.yield_curve, mortality=fix.mortality,
            horizon_age=fix.horizon_age, spread=0.0, valuation_year=2025,
            expenses=fix.expenses, yield_mode_label="flat",
            mortality_mode_label="synthetic", expense_mode_label="manual",
            expense_annual_inflation=0.0,
        )
        return build_myga_xl.build_myga_workbook_from_spec(spec)
    if fix.product_type is ProductType.FIA:
        spec = build_fia_xl.fia_excel_spec_from_launcher(
            contract=fix.contract, yield_curve=fix.yield_curve, mortality=fix.mortality,
            horizon_age=fix.horizon_age, spread=0.0, valuation_year=2025,
            expenses=fix.expenses, yield_mode_label="flat",
            mortality_mode_label="synthetic", expense_mode_label="manual",
            index_s0=100.0,
            index_levels_at_payment=res.index_level_at_payment,
            expense_annual_inflation=0.0,
        )
        return build_fia_xl.build_fia_workbook_from_spec(spec)
    if fix.product_type is ProductType.VARIABLE_ANNUITY:
        spec = build_va_xl.va_excel_spec_from_launcher(
            contract=fix.contract, yield_curve=fix.yield_curve, mortality=fix.mortality,
            horizon_age=fix.horizon_age, spread=0.0, valuation_year=2025,
            expenses=fix.expenses, yield_mode_label="flat",
            mortality_mode_label="synthetic", expense_mode_label="manual",
            index_s0=100.0,
            index_levels_at_payment=res.index_level_at_payment,
            expense_annual_inflation=0.0,
        )
        return build_va_xl.build_va_workbook_from_spec(spec)
    if fix.product_type is ProductType.WHOLE_LIFE:
        spec = build_wl_xl.wl_excel_spec_from_launcher(
            contract=fix.contract, yield_curve=fix.yield_curve, mortality=fix.mortality,
            horizon_age=fix.horizon_age, spread=0.0, valuation_year=2025,
            expenses=fix.expenses, yield_mode_label="flat",
            mortality_mode_label="synthetic", expense_mode_label="manual",
            expense_annual_inflation=0.0,
        )
        return build_wl_xl.build_wl_workbook_from_spec(spec)
    if fix.product_type is ProductType.UNIVERSAL_LIFE:
        spec = build_ul_xl.ul_excel_spec_from_launcher(
            contract=fix.contract, yield_curve=fix.yield_curve, mortality=fix.mortality,
            horizon_age=fix.horizon_age, spread=0.0, valuation_year=2025,
            expenses=fix.expenses, yield_mode_label="flat",
            mortality_mode_label="synthetic", expense_mode_label="manual",
            expense_annual_inflation=0.0,
        )
        return build_ul_xl.build_ul_workbook_from_spec(spec)
    if fix.product_type is ProductType.INDEXED_UL:
        spec = build_iul_xl.iul_excel_spec_from_launcher(
            contract=fix.contract, yield_curve=fix.yield_curve, mortality=fix.mortality,
            horizon_age=fix.horizon_age, spread=0.0, valuation_year=2025,
            expenses=fix.expenses, yield_mode_label="flat",
            mortality_mode_label="synthetic", expense_mode_label="manual",
            index_s0=100.0,
            index_levels_at_payment=res.index_level_at_payment,
            expense_annual_inflation=0.0,
        )
        return build_iul_xl.build_iul_workbook_from_spec(spec)
    if fix.product_type is ProductType.VARIABLE_UL:
        spec = build_vul_xl.vul_excel_spec_from_launcher(
            contract=fix.contract, yield_curve=fix.yield_curve, mortality=fix.mortality,
            horizon_age=fix.horizon_age, spread=0.0, valuation_year=2025,
            expenses=fix.expenses, yield_mode_label="flat",
            mortality_mode_label="synthetic", expense_mode_label="manual",
            index_s0=100.0,
            index_levels_at_payment=res.index_level_at_payment,
            expense_annual_inflation=0.0,
        )
        return build_vul_xl.build_vul_workbook_from_spec(spec)
    raise AssertionError(f"unhandled product {fix.product_type!r}")


def _excel_validates(fix: ProductFixture) -> None:
    """Round-trip the built workbook through the static validator."""
    import io

    import openpyxl

    raw = _build_excel_workbook(fix)
    wb = openpyxl.load_workbook(io.BytesIO(raw))
    excel_workbook_validator.validate_workbook_or_raise(wb)


def _metric_formatter(fix: ProductFixture) -> tuple:
    if fix._result is None:
        _price_deterministic(fix)
    fmt = _PRICING_METRIC_FORMATTERS[fix.product_type]
    metrics = fmt(fix._result)
    assert metrics, f"metric formatter for {fix.product_type!r} returned no metrics"
    return metrics


def _validate_run_inputs(fix: ProductFixture) -> None:
    """Smoke: a happy-path state must produce zero validation errors."""
    state = {
        "issue_age": fix.contract.issue_age,
        "horizon_age": fix.horizon_age,
    }
    errors = validate_run_inputs(fix.product_type, state)
    assert errors == (), (
        f"validate_run_inputs returned {errors!r} for happy-path "
        f"{fix.product_type.value!r} state {state!r}; if a new validator "
        "rejects this baseline, update _validate_run_inputs's state dict "
        "rather than skipping the cell."
    )


SURFACE_RUNNERS: dict[str, Callable[[ProductFixture], Any]] = {
    "pricing_deterministic": _price_deterministic,
    "pricing_monte_carlo": _price_monte_carlo,
    "liability_path": _liability_path,
    "alm_projection": _alm_projection,
    "excel_workbook_build": _build_excel_workbook,
    "excel_workbook_validates": _excel_validates,
    "metric_formatter": _metric_formatter,
    "validate_run_inputs": _validate_run_inputs,
}

SURFACES: tuple[str, ...] = tuple(SURFACE_RUNNERS)
SCENARIOS: tuple[str, ...] = ("baseline", "short_horizon")

# Documented opt-outs. Every entry MUST have a human-readable reason.
# CI / agents reading this file should treat this dict as authoritative:
# silently skipping a cell without an entry here is a bug.
EXPECTED_SKIPS: dict[tuple[ProductType, str, str], str] = {
    # Term Life is currently a deterministic pricer; there is no MC entry
    # point to call. If/when Term grows a Monte Carlo surface, remove this
    # entry rather than adding a new test file.
    (ProductType.TERM_LIFE, "baseline", "pricing_monte_carlo"): (
        "term_projection.py exposes no Monte Carlo entry point yet."
    ),
    (ProductType.TERM_LIFE, "short_horizon", "pricing_monte_carlo"): (
        "term_projection.py exposes no Monte Carlo entry point yet."
    ),
    # MYGA pricing is deterministic given the contract -- no stochastic
    # input. v1 does not expose a Monte Carlo entry point.
    (ProductType.MYGA, "baseline", "pricing_monte_carlo"): (
        "myga_projection.py is deterministic; no Monte Carlo entry point in v1."
    ),
    (ProductType.MYGA, "short_horizon", "pricing_monte_carlo"): (
        "myga_projection.py is deterministic; no Monte Carlo entry point in v1."
    ),
    # WL is deterministic -- mortality is the only stochastic source and v1
    # uses the static CSO 2017 placeholder.
    (ProductType.WHOLE_LIFE, "baseline", "pricing_monte_carlo"): (
        "wl_projection.py is deterministic in v1; no Monte Carlo entry point."
    ),
    (ProductType.WHOLE_LIFE, "short_horizon", "pricing_monte_carlo"): (
        "wl_projection.py is deterministic in v1; no Monte Carlo entry point."
    ),
    # UL has a flat declared rate; no stochastic input drives MC in v1.
    (ProductType.UNIVERSAL_LIFE, "baseline", "pricing_monte_carlo"): (
        "ul_projection.py uses flat declared rate; no Monte Carlo entry point in v1."
    ),
    (ProductType.UNIVERSAL_LIFE, "short_horizon", "pricing_monte_carlo"): (
        "ul_projection.py uses flat declared rate; no Monte Carlo entry point in v1."
    ),
}


# ---------------------------------------------------------------------------
# The matrix.
# ---------------------------------------------------------------------------


def _all_cells() -> list[tuple[ProductType, str, str]]:
    cells: list[tuple[ProductType, str, str]] = []
    for product_type in implemented_product_types():
        if product_type not in _FIXTURE_BUILDERS:
            # If a new ProductType is added but no fixture builder exists, we
            # want the matrix to fail loudly rather than silently shrink.
            raise RuntimeError(
                f"product {product_type!r} is implemented but has no "
                f"_build_<name>_fixture in tests/test_regression_matrix.py; "
                "add one before merging the new product."
            )
        for scenario in SCENARIOS:
            for surface in SURFACES:
                cells.append((product_type, scenario, surface))
    return cells


def _cell_id(cell: tuple[ProductType, str, str]) -> str:
    product, scenario, surface = cell
    return f"{product.value}-{scenario}-{surface}"


@pytest.mark.parametrize("cell", _all_cells(), ids=_cell_id)
def test_regression_matrix_cell(cell: tuple[ProductType, str, str]) -> None:
    """Every (product, scenario, surface) cell either runs or is in EXPECTED_SKIPS."""
    product, scenario, surface = cell
    if cell in EXPECTED_SKIPS:
        pytest.skip(EXPECTED_SKIPS[cell])

    fixture = _FIXTURE_BUILDERS[product](scenario)
    runner = SURFACE_RUNNERS[surface]
    runner(fixture)


# ---------------------------------------------------------------------------
# Meta-checks on the matrix itself.
# ---------------------------------------------------------------------------


def test_every_implemented_product_has_a_fixture_builder() -> None:
    """Every entry in implemented_product_types() must have a fixture builder."""
    missing = [p for p in implemented_product_types() if p not in _FIXTURE_BUILDERS]
    assert not missing, (
        f"products implemented but missing fixture builders: {missing!r}. "
        "Add a `_build_<name>_fixture(scenario)` function above."
    )


def test_every_surface_runner_is_wired() -> None:
    """Every surface name in SURFACES must have a runner in SURFACE_RUNNERS."""
    missing = [s for s in SURFACES if s not in SURFACE_RUNNERS]
    assert not missing, (
        f"surfaces missing runners: {missing!r}. "
        "Add a `_<surface>` function and an entry in SURFACE_RUNNERS."
    )


def test_expected_skips_entries_are_real_cells() -> None:
    """Every EXPECTED_SKIPS key must reference a cell that actually exists.

    This catches stale skips left behind after a surface or scenario is
    renamed, which would otherwise silently pass a no-op test forever.
    """
    real_cells = set(_all_cells())
    stale = [k for k in EXPECTED_SKIPS if k not in real_cells]
    assert not stale, (
        f"EXPECTED_SKIPS contains stale cells (not in the matrix): {stale!r}. "
        "Remove or update them."
    )


def test_expected_skip_reasons_are_nontrivial() -> None:
    """Every EXPECTED_SKIPS reason must be at least 20 chars, no placeholders."""
    bad: list[tuple[tuple[ProductType, str, str], str]] = []
    for cell, reason in EXPECTED_SKIPS.items():
        if len(reason) < 20:
            bad.append((cell, reason))
            continue
        if any(token in reason.lower() for token in ("todo", "fixme", "xxx", "tbd")):
            bad.append((cell, reason))
    assert not bad, (
        f"EXPECTED_SKIPS entries with empty / placeholder reasons: {bad!r}. "
        "Document why the cell is skipped or remove the entry."
    )
