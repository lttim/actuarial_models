"""Portfolio run types: multiple policies, shared scenario, aggregated outputs."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass

import numpy as np

from annuity_model import pricing_projection as sp
from annuity_model.product_registry import ProductContract, ProductType


@dataclass(frozen=True)
class PolicyInput:
    """One policy row in a portfolio run."""

    product_type: ProductType
    contract: ProductContract
    policy_id: str | None = None


@dataclass(frozen=True)
class RunScenario:
    """Economic + expense package shared across all policies in a run."""

    yield_curve: sp.YieldCurve
    mortality: sp.MortalityTableQx | sp.MortalityTableRP2014MP2016
    horizon_age: int
    spread: float
    valuation_year: int | None
    expenses: sp.ExpenseAssumptions | None
    expenses_csv_path: str
    index_scenario_csv_path: str | None
    expense_annual_inflation: float


@dataclass(frozen=True)
class Portfolio:
    policies: tuple[PolicyInput, ...]


@dataclass(frozen=True)
class PolicyResult:
    """Seriatim pricing output for one policy."""

    policy_id: str
    product_type: ProductType
    contract: ProductContract
    pricing: object


@dataclass(frozen=True)
class ProductTypeRollupScalars:
    """Per-:class:`ProductType` scalar rollups (heterogeneous fields are optional)."""

    policy_count: int
    sum_single_premium: float | None = None
    sum_undiscounted_cashflows: float | None = None


@dataclass(frozen=True)
class PortfolioResult:
    """Full portfolio run: seriatim, by-type paths, total path, optional ALM."""

    policy_results: tuple[PolicyResult, ...]
    rollups_by_product_type: Mapping[ProductType, sp.LiabilityPath]
    product_type_scalar_rollups: Mapping[ProductType, ProductTypeRollupScalars]
    liability_path_total: sp.LiabilityPath
    alm_result: sp.ALMResult | None


def default_policy_id(index: int) -> str:
    return f"p{index}"


def compute_scalar_rollups_for_type(results: Sequence[PolicyResult]) -> ProductTypeRollupScalars:
    """Derive scalar rollups from seriatim :class:`PolicyResult` for one product type."""
    n = len(results)
    premiums: list[float] = []
    cf_sums: list[float] = []
    for pr in results:
        p = pr.pricing
        if hasattr(p, "single_premium"):
            premiums.append(float(p.single_premium))
        if hasattr(p, "expected_total_cashflows"):
            cf_sums.append(float(np.sum(np.asarray(p.expected_total_cashflows, dtype=float))))
    return ProductTypeRollupScalars(
        policy_count=n,
        sum_single_premium=float(sum(premiums)) if len(premiums) == n else None,
        sum_undiscounted_cashflows=float(sum(cf_sums)) if len(cf_sums) == n else None,
    )


__all__ = [
    "PolicyInput",
    "PolicyResult",
    "Portfolio",
    "PortfolioResult",
    "ProductTypeRollupScalars",
    "RunScenario",
    "compute_scalar_rollups_for_type",
    "default_policy_id",
]
