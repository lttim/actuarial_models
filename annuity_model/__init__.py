"""annuity_model -- actuarial pricing & ALM platform.

This package exposes the public surface used by the Streamlit UI, the test
suite, and any external integrations. Internal modules may still be imported
by name (`import pricing_projection as sp`) until the planned `src/` layout
ships, but new code SHOULD prefer the symbols re-exported here so that future
module reshuffles do not break callers.

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

import sys
from pathlib import Path
from types import ModuleType

# Transitional import bootstrap ------------------------------------------------
#
# The runtime still contains legacy flat imports such as ``import pricing_projection``.
# When this package is imported from the repository root, those sibling modules
# are not on ``sys.path`` by default.  Add this package directory explicitly so
# the existing launchers, tests, and Streamlit entrypoints keep their current
# behavior until the larger src-layout migration can rewrite the imports.
_PACKAGE_DIR = Path(__file__).resolve().parent
_PACKAGE_DIR_STR = str(_PACKAGE_DIR)
if _PACKAGE_DIR_STR not in sys.path:
    sys.path.insert(0, _PACKAGE_DIR_STR)


def _alias_legacy_module(flat_name: str) -> None:
    """Expose a loaded legacy flat module under ``annuity_model.<name>`` too."""
    module = sys.modules.get(flat_name)
    if isinstance(module, ModuleType):
        sys.modules.setdefault(f"{__name__}.{flat_name}", module)


def _alias_legacy_prefix(flat_prefix: str) -> None:
    """Expose loaded legacy subpackage modules under the package namespace."""
    package_prefix = f"{__name__}.{flat_prefix}"
    for loaded_name, module in list(sys.modules.items()):
        if loaded_name == flat_prefix or loaded_name.startswith(f"{flat_prefix}."):
            suffix = loaded_name[len(flat_prefix) :]
            sys.modules.setdefault(f"{package_prefix}{suffix}", module)


# Logging -----------------------------------------------------------------
from _logging import configure_logging, get_logger

# Excel pipeline -----------------------------------------------------------
from excel_workbook_validator import (
    ExcelWorkbookValidationError,
    FormulaIssue,
    validate_formula,
    validate_workbook,
    validate_workbook_or_raise,
)

# Portfolio (multi-policy) --------------------------------------------------
from liability_aggregation import (
    aggregate_by_product_type,
    aggregate_liability_paths,
)
from liability_layouts import (
    LIABILITY_LAYOUTS,
    LiabilityLayout,
    liability_layout_for,
)
from portfolio import (
    PolicyInput,
    Portfolio,
    PortfolioResult,
    RunScenario,
)
from portfolio_runner import run_portfolio

# Engine surface -----------------------------------------------------------
from pricing_projection import (
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
from product_excel import build_product_workbook

# Product registry --------------------------------------------------------
from product_registry import (
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
from products import (
    ProductDefinition,
    discover_products,
    get_product_definition,
    iter_product_definitions,
)
from rila_projection import (
    RILAContract,
    liability_path_from_rila_projection,
    price_rila_single_premium,
)
from term_projection import (
    TermLifeContract,
    liability_path_from_term_projection,
    price_term_life_level_monthly,
)

for _flat_module_name in (
    "_logging",
    "excel_workbook_validator",
    "liability_aggregation",
    "liability_layouts",
    "portfolio",
    "portfolio_runner",
    "pricing_projection",
    "product_excel",
    "product_registry",
    "rila_projection",
    "term_projection",
):
    _alias_legacy_module(_flat_module_name)
_alias_legacy_prefix("products")
del _flat_module_name

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
