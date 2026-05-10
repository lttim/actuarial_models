"""Term Life product subpackage.

See :mod:`products.spia` for the pattern; this module is the Term-Life
analog.
"""

from __future__ import annotations

from annuity_model.product_excel import _BUILDER_REGISTRY
from annuity_model.product_registry import ProductCapabilities, ProductType
from annuity_model.products import ProductDefinition, register_product
from annuity_model.products.term.engine import (
    TermLifeContract,
    TermLifeProjectionResult,
    liability_path_from_term_projection,
    price_term_life_level_monthly,
)
from annuity_model.products.term.excel import TermExcelBuildSpec, term_excel_spec_from_launcher
from annuity_model.products.term.ui import TERM_ADAPTER, term_metric_formatter, term_ui_config

DEFINITION = register_product(
    ProductDefinition(
        product_type=ProductType.TERM_LIFE,
        display_name="Term Life (20Y)",
        contract_type=TermLifeContract,
        result_type=TermLifeProjectionResult,
        builder_spec_type=TermExcelBuildSpec,
        adapter=TERM_ADAPTER,
        builder=_BUILDER_REGISTRY[ProductType.TERM_LIFE],
        liability_path_converter=liability_path_from_term_projection,
        metric_formatter=term_metric_formatter,
        capabilities=ProductCapabilities(
            supports_economic_scenario=False, supports_monte_carlo=False
        ),
        ui_config=term_ui_config,
        mortality_mode_options=("us_ssa_2015_period", "qx_csv", "synthetic"),
        default_mortality_mode="us_ssa_2015_period",
        validator=None,
        order=1,
    )
)


__all__ = [
    "DEFINITION",
    "TERM_ADAPTER",
    "TermExcelBuildSpec",
    "TermLifeContract",
    "TermLifeProjectionResult",
    "liability_path_from_term_projection",
    "price_term_life_level_monthly",
    "term_excel_spec_from_launcher",
    "term_metric_formatter",
    "term_ui_config",
]
