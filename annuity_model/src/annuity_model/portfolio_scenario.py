"""Default :class:`portfolio.RunScenario` for portfolio CLI / tests (Pricing Run first-paint economics)."""

from __future__ import annotations

from annuity_model.portfolio import RunScenario
from annuity_model.pricing_scenario_materialize import cli_default_run_scenario
from annuity_model.product_registry import ProductType


def default_run_scenario(
    *,
    default_product_type: ProductType = ProductType.SPIA,
) -> RunScenario:
    """Deterministic scenario aligned with ``build_run_form_seed_defaults`` / Pricing Run."""
    return cli_default_run_scenario(default_product_type=default_product_type)


__all__ = ["default_run_scenario"]
