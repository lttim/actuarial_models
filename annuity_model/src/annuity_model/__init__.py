"""annuity_model -- actuarial pricing & ALM platform.

This package exposes the public surface used by the Streamlit UI, the test
suite, and any external integrations. Runtime modules live under the standard
``src/annuity_model`` package layout; use package-qualified imports such as
``annuity_model.pricing_projection`` instead of legacy flat module names.

Quick reference:

* Engines: :class:`SPIAContract`, :class:`TermLifeContract`, :class:`RILAContract`
* Pricing entry points: :func:`price_spia_single_premium`,
  :func:`price_term_life_level_monthly`, :func:`price_rila_single_premium`
* ALM: :func:`run_alm_projection`, :func:`run_alm_projection_from_pricing_result`
* Excel pipeline: :func:`build_product_workbook`, :func:`validate_workbook_or_raise`
* Product registry: :class:`ProductType`, :class:`ProductAdapter`,
  :func:`get_product_adapter`, :data:`LIABILITY_LAYOUTS`
"""

from __future__ import annotations

# Logging -----------------------------------------------------------------
from annuity_model._logging import configure_logging, get_logger

# Excel pipeline -----------------------------------------------------------
from annuity_model.excel_workbook_validator import (
    ExcelWorkbookValidationError,
    FormulaIssue,
    validate_formula,
    validate_workbook,
    validate_workbook_or_raise,
)

# Portfolio (multi-policy) --------------------------------------------------
from annuity_model.liability_aggregation import (
    aggregate_by_product_type,
    aggregate_liability_paths,
)
from annuity_model.liability_layouts import (
    LIABILITY_LAYOUTS,
    LiabilityLayout,
    liability_layout_for,
)
from annuity_model.portfolio import (
    PolicyInput,
    Portfolio,
    PortfolioResult,
    RunScenario,
)
from annuity_model.portfolio_runner import run_portfolio

# Engine surface -----------------------------------------------------------
from annuity_model.pricing_projection import (
    ALMAllocationSpec,
    ALMAssumptions,
    ALMBucketSpec,
    ExpenseAssumptions,
    MortalityTableQx,
    SPIAContract,
    YieldCurve,
    alm_default_allocation_spec,
    liability_path_from_spia_projection,
    price_spia_single_premium,
    run_alm_projection,
    run_alm_projection_from_liability_path,
    run_alm_projection_from_pricing_result,
)
from annuity_model.product_excel import build_product_workbook

# Product registry --------------------------------------------------------
from annuity_model.product_registry import (
    ProductAdapter,
    ProductContract,
    ProductType,
    get_pricing_metrics,
    get_product_adapter,
    get_product_capabilities,
    product_label,
    product_options_for_ui,
    validate_run_inputs,
)

# Unified per-product definitions (P1, 2026-04). Consolidates the five
# legacy per-product registries into one immutable ProductDefinition per
# product. Auto-discovers shims under annuity_model/products/. See
# annuity_model/products/__init__.py for the rationale.
from annuity_model.products import (
    ProductDefinition,
    discover_products,
    get_product_definition,
    iter_product_definitions,
)
from annuity_model.rila_projection import (
    RILAContract,
    liability_path_from_rila_projection,
    price_rila_single_premium,
)
from annuity_model.term_projection import (
    TermLifeContract,
    liability_path_from_term_projection,
    price_term_life_level_monthly,
)

__all__ = [
    "PolicyInput",
    "Portfolio",
    "PortfolioResult",
    "RunScenario",
    "aggregate_by_product_type",
    "aggregate_liability_paths",
    "run_portfolio",
    "ALMAllocationSpec",
    "ALMAssumptions",
    "ALMBucketSpec",
    "ExcelWorkbookValidationError",
    "ExpenseAssumptions",
    "FormulaIssue",
    "LIABILITY_LAYOUTS",
    "LiabilityLayout",
    "MortalityTableQx",
    "ProductAdapter",
    "ProductContract",
    "ProductDefinition",
    "ProductType",
    "RILAContract",
    "SPIAContract",
    "TermLifeContract",
    "YieldCurve",
    "alm_default_allocation_spec",
    "build_product_workbook",
    "configure_logging",
    "discover_products",
    "get_logger",
    "get_pricing_metrics",
    "get_product_adapter",
    "get_product_capabilities",
    "get_product_definition",
    "iter_product_definitions",
    "liability_layout_for",
    "liability_path_from_rila_projection",
    "liability_path_from_spia_projection",
    "liability_path_from_term_projection",
    "price_rila_single_premium",
    "price_spia_single_premium",
    "price_term_life_level_monthly",
    "product_label",
    "product_options_for_ui",
    "run_alm_projection",
    "run_alm_projection_from_liability_path",
    "run_alm_projection_from_pricing_result",
    "validate_formula",
    "validate_run_inputs",
    "validate_workbook",
    "validate_workbook_or_raise",
]
