"""SPIA product subpackage.

Forward-compatible import surface for the SPIA product. The actual
implementation still lives in the legacy flat modules
(``pricing_projection``, ``build_pricing_excel_workbook``,
``product_registry``); this subpackage exposes them under the canonical
``products.spia.{schema, engine, excel, ui}`` path that new code SHOULD
prefer.

Layout (mirrored across :mod:`products.term` and :mod:`products.rila`):

* :mod:`products.spia.schema`  -- contract dataclass(es), result classes
* :mod:`products.spia.engine`  -- pricing + ALM converter callables
* :mod:`products.spia.excel`   -- Excel build-spec + builder
* :mod:`products.spia.ui`      -- adapter + UI config

The submodules are re-export shims today, not new implementations. A
later wave can move the implementation here without touching any caller
that already imports from ``products.spia.*`` -- that is the whole point
of having the shim land first.

The :class:`~products.ProductDefinition` for SPIA is built and registered
at the bottom of this file; importing :mod:`products.spia` is sufficient
to make ``get_product_definition(ProductType.SPIA)`` resolve.
"""

from __future__ import annotations

from product_excel import _BUILDER_REGISTRY
from product_registry import ProductType, product_label
from products import ProductDefinition, register_product
from products.spia.engine import (
    SPIAContract,
    SPIAProjectionResult,
    liability_path_from_spia_projection,
    price_spia_single_premium,
)
from products.spia.excel import ExcelBuildSpec, build_workbook_from_spec
from products.spia.ui import SPIA_ADAPTER, spia_metric_formatter, spia_ui_config

DEFINITION = register_product(
    ProductDefinition(
        product_type=ProductType.SPIA,
        display_name=product_label(ProductType.SPIA),
        contract_type=SPIAContract,
        result_type=SPIAProjectionResult,
        builder_spec_type=ExcelBuildSpec,
        adapter=SPIA_ADAPTER,
        builder=_BUILDER_REGISTRY[ProductType.SPIA],
        liability_path_converter=liability_path_from_spia_projection,
        metric_formatter=spia_metric_formatter,
    )
)


__all__ = [
    "DEFINITION",
    "ExcelBuildSpec",
    "SPIAContract",
    "SPIAProjectionResult",
    "SPIA_ADAPTER",
    "build_workbook_from_spec",
    "liability_path_from_spia_projection",
    "price_spia_single_premium",
    "spia_metric_formatter",
    "spia_ui_config",
]
