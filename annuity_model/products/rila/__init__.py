"""RILA product subpackage.

See :mod:`products.spia` for the pattern; this module is the RILA analog.
"""

from __future__ import annotations

from product_excel import _BUILDER_REGISTRY
from product_registry import ProductType, product_label
from products import ProductDefinition, register_product
from products.rila.engine import (
    RILAContract,
    RILAProjectionResult,
    liability_path_from_rila_projection,
    price_rila_single_premium,
)
from products.rila.excel import RILAExcelBuildSpec, rila_excel_spec_from_launcher
from products.rila.ui import RILA_ADAPTER, rila_metric_formatter, rila_ui_config

DEFINITION = register_product(
    ProductDefinition(
        product_type=ProductType.RILA,
        display_name=product_label(ProductType.RILA),
        contract_type=RILAContract,
        result_type=RILAProjectionResult,
        builder_spec_type=RILAExcelBuildSpec,
        adapter=RILA_ADAPTER,
        builder=_BUILDER_REGISTRY[ProductType.RILA],
        liability_path_converter=liability_path_from_rila_projection,
        metric_formatter=rila_metric_formatter,
    )
)


__all__ = [
    "DEFINITION",
    "RILAContract",
    "RILAExcelBuildSpec",
    "RILAProjectionResult",
    "RILA_ADAPTER",
    "liability_path_from_rila_projection",
    "price_rila_single_premium",
    "rila_excel_spec_from_launcher",
    "rila_metric_formatter",
    "rila_ui_config",
]
