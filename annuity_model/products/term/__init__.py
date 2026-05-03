"""Term Life product subpackage.

See :mod:`products.spia` for the pattern; this module is the Term-Life
analog.
"""

from __future__ import annotations

from product_excel import _BUILDER_REGISTRY
from product_registry import ProductType, product_label
from products import ProductDefinition, register_product
from products.term.engine import (
    TermLifeContract,
    TermLifeProjectionResult,
    liability_path_from_term_projection,
    price_term_life_level_monthly,
)
from products.term.excel import TermExcelBuildSpec, term_excel_spec_from_launcher
from products.term.ui import TERM_ADAPTER, term_metric_formatter, term_ui_config

DEFINITION = register_product(
    ProductDefinition(
        product_type=ProductType.TERM_LIFE,
        display_name=product_label(ProductType.TERM_LIFE),
        contract_type=TermLifeContract,
        result_type=TermLifeProjectionResult,
        builder_spec_type=TermExcelBuildSpec,
        adapter=TERM_ADAPTER,
        builder=_BUILDER_REGISTRY[ProductType.TERM_LIFE],
        liability_path_converter=liability_path_from_term_projection,
        metric_formatter=term_metric_formatter,
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
