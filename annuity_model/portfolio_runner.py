"""Run pricing for a :class:`portfolio.Portfolio` and aggregate liability paths."""

from __future__ import annotations

from collections import defaultdict
from concurrent.futures import ProcessPoolExecutor

import pricing_projection as sp
from _observability import traced
from liability_aggregation import (
    aggregate_by_product_type,
    aggregate_liability_paths,
    assert_rollups_sum_to_total,
)
from liability_dispatch import liability_path_for
from parity_constants import PORTFOLIO_ROLLUP_TOL
from portfolio import (
    PolicyInput,
    PolicyResult,
    Portfolio,
    PortfolioResult,
    ProductTypeRollupScalars,
    RunScenario,
    compute_scalar_rollups_for_type,
    default_policy_id,
)
from product_registry import ProductAdapter, ProductType, get_product_adapter


@traced("portfolio.adapter_price")
def _adapter_price(adapter: ProductAdapter, pol: PolicyInput, scenario: RunScenario) -> object:
    return adapter.price(
        contract=pol.contract,
        yield_curve=scenario.yield_curve,
        mortality=scenario.mortality,
        horizon_age=scenario.horizon_age,
        spread=scenario.spread,
        valuation_year=scenario.valuation_year,
        expenses=scenario.expenses,
        expenses_csv_path=scenario.expenses_csv_path,
        index_scenario_csv_path=scenario.index_scenario_csv_path,
        expense_annual_inflation=scenario.expense_annual_inflation,
    )


def _worker_pack(args: tuple[int, PolicyInput, RunScenario]) -> tuple[PolicyResult, tuple[ProductType, sp.LiabilityPath]]:
    """Picklable worker for :class:`ProcessPoolExecutor` (one policy)."""
    i, pol, scenario = args
    adapter = get_product_adapter(pol.product_type)
    pricing = _adapter_price(adapter, pol, scenario)
    pid = pol.policy_id if pol.policy_id is not None else default_policy_id(i)
    pr = PolicyResult(
        policy_id=pid,
        product_type=pol.product_type,
        contract=pol.contract,
        pricing=pricing,
    )
    return pr, (pol.product_type, liability_path_for(pricing))


def _initial_asset_market_value(policy_results: tuple[PolicyResult, ...]) -> float:
    """Sum ``single_premium`` when every result exposes it; else raise."""
    total = 0.0
    for pr in policy_results:
        p = pr.pricing
        if not hasattr(p, "single_premium"):
            raise ValueError(
                "initial_asset_market_value cannot be inferred: "
                f"{type(p).__name__} has no single_premium; pass explicit AUM to run_portfolio."
            )
        total += float(p.single_premium)
    if total <= 0.0:
        raise ValueError("initial_asset_market_value inferred from premiums must be > 0.")
    return total


def run_portfolio(
    *,
    portfolio: Portfolio,
    scenario: RunScenario,
    alm_assumptions: sp.ALMAssumptions | None = None,
    initial_asset_market_value: float | None = None,
    max_workers: int = 1,
) -> PortfolioResult:
    """Price every policy, build by-type and total liability paths, optional ALM.

    ``max_workers`` controls parallel pricing via :class:`ProcessPoolExecutor`.
    The default ``1`` is deterministic single-process; values ``>1`` split work
    across processes (each policy must pickle cleanly).
    """
    packed: list[tuple[int, PolicyInput, RunScenario]] = [
        (i, pol, scenario) for i, pol in enumerate(portfolio.policies)
    ]
    if not packed:
        raise ValueError("run_portfolio requires at least one policy.")

    if max_workers <= 1:
        rows = [_worker_pack(t) for t in packed]
    else:
        workers = min(int(max_workers), len(packed))
        with ProcessPoolExecutor(max_workers=workers) as ex:
            rows = list(ex.map(_worker_pack, packed))

    policy_results = [r[0] for r in rows]
    typed_paths = [r[1] for r in rows]

    liability_total = aggregate_liability_paths([p for _, p in typed_paths])
    rollups = aggregate_by_product_type(typed_paths)
    assert_rollups_sum_to_total(
        rollups_by_product_type=rollups,
        portfolio=liability_total,
        rtol=0.0,
        atol=PORTFOLIO_ROLLUP_TOL,
    )

    by_type_results: dict[ProductType, list[PolicyResult]] = defaultdict(list)
    for pr in policy_results:
        by_type_results[pr.product_type].append(pr)

    scalar_rollups: dict[ProductType, ProductTypeRollupScalars] = {}
    for pt in sorted(rollups, key=lambda x: x.value):
        scalar_rollups[pt] = compute_scalar_rollups_for_type(tuple(by_type_results[pt]))

    pr_tuple = tuple(policy_results)
    alm_out: sp.ALMResult | None = None
    if alm_assumptions is not None:
        aum = (
            float(initial_asset_market_value)
            if initial_asset_market_value is not None
            else _initial_asset_market_value(pr_tuple)
        )
        alm_out = sp.run_alm_projection_from_liability_path(
            liability_path=liability_total,
            yield_curve=scenario.yield_curve,
            spread=scenario.spread,
            assumptions=alm_assumptions,
            initial_asset_market_value=aum,
        )

    return PortfolioResult(
        policy_results=pr_tuple,
        rollups_by_product_type=rollups,
        product_type_scalar_rollups=scalar_rollups,
        liability_path_total=liability_total,
        alm_result=alm_out,
    )


__all__ = ["run_portfolio"]
