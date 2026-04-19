"""Symbol-identity guard for the per-product subpackage shims.

The :mod:`products.spia`, :mod:`products.term`, and :mod:`products.rila`
subpackages are *re-export shims* over the legacy flat modules. New code
SHOULD import from the canonical ``products.<name>.{schema,engine,excel,ui}``
paths, but until the implementation physically moves into the subpackages
the two import paths must resolve to the **same object** (``is`` check).

If a contributor accidentally creates a parallel implementation under
``products/<name>/engine.py`` -- e.g. by copying the dataclass instead of
re-exporting it -- the codebase would silently ship two SPIA pricers
that drift apart over time. These identity assertions catch that at PR
time.

When the implementation eventually moves into the subpackages, this test
gets inverted: the legacy module becomes the shim and re-exports from
``products.<name>``. The shape of the assertions stays the same; only
the direction changes.
"""

from __future__ import annotations

import pytest

import build_pricing_excel_workbook as bpw
import build_rila_excel_workbook as brw
import build_term_excel_workbook as btw
import pricing_projection as sp
import rila_projection as rp
import term_projection as tp

import products.rila as rila_pkg
import products.rila.engine as rila_engine
import products.rila.excel as rila_excel
import products.rila.schema as rila_schema
import products.rila.ui as rila_ui
import products.spia as spia_pkg
import products.spia.engine as spia_engine
import products.spia.excel as spia_excel
import products.spia.schema as spia_schema
import products.spia.ui as spia_ui
import products.term as term_pkg
import products.term.engine as term_engine
import products.term.excel as term_excel
import products.term.schema as term_schema
import products.term.ui as term_ui
from product_registry import ProductType, get_product_adapter

pytestmark = [pytest.mark.invariant]


# ---------------------------------------------------------------------------
# SPIA
# ---------------------------------------------------------------------------


def test_spia_schema_reexports_legacy_classes() -> None:
    assert spia_schema.SPIAContract is sp.SPIAContract
    assert spia_schema.SPIAProjectionResult is sp.SPIAProjectionResult
    assert spia_schema.SPIAMonteCarloResult is sp.SPIAMonteCarloResult


def test_spia_engine_reexports_legacy_callables() -> None:
    assert spia_engine.price_spia_single_premium is sp.price_spia_single_premium
    assert (
        spia_engine.price_spia_single_premium_monte_carlo
        is sp.price_spia_single_premium_monte_carlo
    )
    assert (
        spia_engine.liability_path_from_spia_projection
        is sp.liability_path_from_spia_projection
    )


def test_spia_excel_reexports_legacy_builder() -> None:
    assert spia_excel.ExcelBuildSpec is bpw.ExcelBuildSpec
    assert spia_excel.excel_spec_from_launcher is bpw.excel_spec_from_launcher
    assert spia_excel.build_workbook_from_spec is bpw.build_workbook_from_spec


def test_spia_ui_reexports_adapter() -> None:
    assert spia_ui.SPIA_ADAPTER is get_product_adapter(ProductType.SPIA)


def test_spia_package_top_level_reexports_match_subpackage() -> None:
    assert spia_pkg.SPIAContract is spia_schema.SPIAContract
    assert spia_pkg.ExcelBuildSpec is spia_excel.ExcelBuildSpec
    assert spia_pkg.SPIA_ADAPTER is spia_ui.SPIA_ADAPTER
    assert spia_pkg.DEFINITION.product_type is ProductType.SPIA


# ---------------------------------------------------------------------------
# Term Life
# ---------------------------------------------------------------------------


def test_term_schema_reexports_legacy_classes() -> None:
    assert term_schema.TermLifeContract is tp.TermLifeContract
    assert term_schema.TermLifeProjectionResult is tp.TermLifeProjectionResult


def test_term_engine_reexports_legacy_callables() -> None:
    assert term_engine.price_term_life_level_monthly is tp.price_term_life_level_monthly
    assert (
        term_engine.liability_path_from_term_projection
        is tp.liability_path_from_term_projection
    )


def test_term_excel_reexports_legacy_builder() -> None:
    assert term_excel.TermExcelBuildSpec is btw.TermExcelBuildSpec
    assert term_excel.term_excel_spec_from_launcher is btw.term_excel_spec_from_launcher
    assert (
        term_excel.build_term_workbook_from_spec is btw.build_term_workbook_from_spec
    )


def test_term_ui_reexports_adapter_and_parsers() -> None:
    assert term_ui.TERM_ADAPTER is get_product_adapter(ProductType.TERM_LIFE)
    # Parsers must reference the SAME callable as product_registry exports;
    # divergence here would silently break the AST guard in
    # tests/test_pricing_ui_term_config.py.
    from product_registry import (
        parse_term_benefit_timing_label,
        parse_term_length_label_to_years,
        parse_term_premium_mode_label,
    )

    assert term_ui.parse_term_length_label_to_years is parse_term_length_label_to_years
    assert term_ui.parse_term_premium_mode_label is parse_term_premium_mode_label
    assert (
        term_ui.parse_term_benefit_timing_label is parse_term_benefit_timing_label
    )


def test_term_package_top_level_reexports_match_subpackage() -> None:
    assert term_pkg.TermLifeContract is term_schema.TermLifeContract
    assert term_pkg.TermExcelBuildSpec is term_excel.TermExcelBuildSpec
    assert term_pkg.TERM_ADAPTER is term_ui.TERM_ADAPTER
    assert term_pkg.DEFINITION.product_type is ProductType.TERM_LIFE


# ---------------------------------------------------------------------------
# RILA
# ---------------------------------------------------------------------------


def test_rila_schema_reexports_legacy_classes() -> None:
    assert rila_schema.RILAContract is rp.RILAContract
    assert rila_schema.RILAProjectionResult is rp.RILAProjectionResult
    assert rila_schema.RILAMonteCarloResult is rp.RILAMonteCarloResult
    assert rila_schema.RILAPricingInfeasibleError is rp.RILAPricingInfeasibleError


def test_rila_engine_reexports_legacy_callables() -> None:
    assert rila_engine.price_rila_single_premium is rp.price_rila_single_premium
    assert (
        rila_engine.price_rila_single_premium_monte_carlo
        is rp.price_rila_single_premium_monte_carlo
    )
    assert (
        rila_engine.liability_path_from_rila_projection
        is rp.liability_path_from_rila_projection
    )


def test_rila_excel_reexports_legacy_builder() -> None:
    assert rila_excel.RILAExcelBuildSpec is brw.RILAExcelBuildSpec
    assert rila_excel.rila_excel_spec_from_launcher is brw.rila_excel_spec_from_launcher
    assert (
        rila_excel.build_rila_workbook_from_spec is brw.build_rila_workbook_from_spec
    )


def test_rila_ui_reexports_adapter() -> None:
    assert rila_ui.RILA_ADAPTER is get_product_adapter(ProductType.RILA)


def test_rila_package_top_level_reexports_match_subpackage() -> None:
    assert rila_pkg.RILAContract is rila_schema.RILAContract
    assert rila_pkg.RILAExcelBuildSpec is rila_excel.RILAExcelBuildSpec
    assert rila_pkg.RILA_ADAPTER is rila_ui.RILA_ADAPTER
    assert rila_pkg.DEFINITION.product_type is ProductType.RILA


# ---------------------------------------------------------------------------
# Cross-product structure: every implemented product must expose the same
# four submodules so contract scaffolding tooling can rely on the layout.
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    "package",
    [spia_pkg, term_pkg, rila_pkg],
    ids=lambda pkg: pkg.__name__,
)
def test_every_product_subpackage_exposes_canonical_layout(package) -> None:
    """schema/engine/excel/ui submodules MUST be importable for every
    implemented product. Scaffolding scripts (P1.5) rely on this layout
    when generating a new product."""
    import importlib

    for submodule in ("schema", "engine", "excel", "ui"):
        mod = importlib.import_module(f"{package.__name__}.{submodule}")
        assert mod is not None, (
            f"{package.__name__}.{submodule} failed to import. The "
            "schema/engine/excel/ui layout is mandatory for every product "
            "subpackage; add the file as a re-export shim."
        )


@pytest.mark.parametrize(
    "package",
    [spia_pkg, term_pkg, rila_pkg],
    ids=lambda pkg: pkg.__name__,
)
def test_every_product_subpackage_publishes_a_definition(package) -> None:
    from products import ProductDefinition

    assert hasattr(package, "DEFINITION"), (
        f"{package.__name__}.DEFINITION is missing. Every product "
        "subpackage MUST register and re-export a ProductDefinition."
    )
    assert isinstance(package.DEFINITION, ProductDefinition)
