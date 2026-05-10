"""RILA product subpackage.

See :mod:`products.spia` for the pattern; this module is the RILA analog.
"""

from __future__ import annotations

from annuity_model.product_excel import _BUILDER_REGISTRY
from annuity_model.product_registry import ProductCapabilities, ProductType
from annuity_model.products import ProductDefinition, register_product
from annuity_model.products.rila.engine import (
    RILAContract,
    RILAProjectionResult,
    liability_path_from_rila_projection,
    price_rila_single_premium,
)
from annuity_model.products.rila.excel import RILAExcelBuildSpec, rila_excel_spec_from_launcher
from annuity_model.products.rila.ui import RILA_ADAPTER, rila_metric_formatter, rila_ui_config

DEFINITION = register_product(
    ProductDefinition(
        product_type=ProductType.RILA,
        display_name="RILA (accumulation)",
        contract_type=RILAContract,
        result_type=RILAProjectionResult,
        builder_spec_type=RILAExcelBuildSpec,
        adapter=RILA_ADAPTER,
        builder=_BUILDER_REGISTRY[ProductType.RILA],
        liability_path_converter=liability_path_from_rila_projection,
        metric_formatter=rila_metric_formatter,
        capabilities=ProductCapabilities(
            supports_economic_scenario=True, supports_monte_carlo=True
        ),
        ui_config=rila_ui_config,
        mortality_mode_options=("synthetic", "qx_csv", "rp2014_mp2016"),
        default_mortality_mode="rp2014_mp2016",
        validator=None,
        order=2,
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
