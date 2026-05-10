from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from types import MappingProxyType
from typing import Any, cast

from annuity_model import pricing_projection as sp
from annuity_model.build_fia_excel_workbook import FIAExcelBuildSpec, build_fia_workbook_from_spec
from annuity_model.build_iul_excel_workbook import IULExcelBuildSpec, build_iul_workbook_from_spec
from annuity_model.build_myga_excel_workbook import (
    MYGAExcelBuildSpec,
    build_myga_workbook_from_spec,
)
from annuity_model.build_pricing_excel_workbook import (
    ALMExcelSnapshot,
    ExcelBuildSpec,
    ExcelPythonSnapshot,
    MCExcelSnapshot,
    build_workbook_from_spec,
)
from annuity_model.build_rila_excel_workbook import (
    RILAExcelBuildSpec,
    build_rila_workbook_from_spec,
)
from annuity_model.build_term_excel_workbook import (
    TermExcelBuildSpec,
    build_term_workbook_from_spec,
)
from annuity_model.build_ul_excel_workbook import ULExcelBuildSpec, build_ul_workbook_from_spec
from annuity_model.build_va_excel_workbook import VAExcelBuildSpec, build_va_workbook_from_spec
from annuity_model.build_vul_excel_workbook import VULExcelBuildSpec, build_vul_workbook_from_spec
from annuity_model.build_wl_excel_workbook import WLExcelBuildSpec, build_wl_workbook_from_spec
from annuity_model.product_registry import ProductType

# Per-product workbook builder. The signature accepts the union of all
# kwargs the dispatcher might pass; individual builders pull what they need
# (spec, out_path, alm_*) and ignore the rest. Returning bytes (the .xlsx
# blob) is required so callers downstream of build_product_workbook can
# write to disk *or* hand off to a Streamlit download_button without going
# through the filesystem.
ProductWorkbookBuilder = Callable[..., bytes]

_BUILDER_REGISTRY: dict[ProductType, ProductWorkbookBuilder] = {}
# Each entry's *expected* spec class. Validated at dispatch time so a
# wrong-type spec fails fast with a clear TypeError instead of a confusing
# AttributeError deep inside the builder. None means "no spec-type
# enforcement" (currently unused; reserved for future products that may
# accept a union spec).
_BUILDER_SPEC_TYPES: dict[ProductType, type | None] = {}


def _workbook_bytes(value: Any) -> bytes:
    """Type-narrow per-product workbook builders that return xlsx bytes."""
    return cast(bytes, value)


def register_builder(
    product_type: ProductType,
    *,
    spec_type: type | None,
) -> Callable[[ProductWorkbookBuilder], ProductWorkbookBuilder]:
    """Register *fn* as the workbook builder for *product_type*.

    Used as a decorator on the per-product wrapper functions defined below.
    Re-registering the same product_type raises RuntimeError -- in this
    codebase that almost always means a copy-paste mistake (two adapters
    claiming the same enum), not an intentional override. If you really
    need to override (e.g. for a test), pop from `_BUILDER_REGISTRY`
    first.

    spec_type is the dataclass that the dispatcher will isinstance-check
    the incoming `spec` against before calling fn. Pass `None` only if the
    builder accepts a union of spec types (none today).
    """

    def decorator(fn: ProductWorkbookBuilder) -> ProductWorkbookBuilder:
        if product_type in _BUILDER_REGISTRY:
            raise RuntimeError(
                f"Workbook builder for {product_type!r} is already registered "
                f"(existing={_BUILDER_REGISTRY[product_type]!r}, new={fn!r}). "
                "Pop from _BUILDER_REGISTRY first if this is intentional."
            )
        _BUILDER_REGISTRY[product_type] = fn
        _BUILDER_SPEC_TYPES[product_type] = spec_type
        return fn

    return decorator


@register_builder(ProductType.SPIA, spec_type=ExcelBuildSpec)
def _build_spia_workbook(
    *,
    spec: Any,
    out_path: str | Path | None,
    python_snapshot: ExcelPythonSnapshot | None,
    mc_snapshot: MCExcelSnapshot | None,
    alm_snapshot: ALMExcelSnapshot | None,
    alm_assumptions: sp.ALMAssumptions | None,
) -> bytes:
    return _workbook_bytes(
        build_workbook_from_spec(
            spec,
            out_path=Path(out_path) if out_path is not None else None,
            python_snapshot=python_snapshot,
            mc_snapshot=mc_snapshot,
            alm_snapshot=alm_snapshot,
            alm_assumptions=alm_assumptions,
        )
    )


@register_builder(ProductType.TERM_LIFE, spec_type=TermExcelBuildSpec)
def _build_term_workbook(
    *,
    spec: Any,
    out_path: str | Path | None,
    python_snapshot: ExcelPythonSnapshot | None,
    mc_snapshot: MCExcelSnapshot | None,
    alm_snapshot: ALMExcelSnapshot | None,
    alm_assumptions: sp.ALMAssumptions | None,
) -> bytes:
    # Term builder doesn't take python_snapshot / mc_snapshot today (no
    # ESG/MC sheets are emitted for level-monthly term). Accept and drop
    # them so the dispatcher signature is uniform.
    del python_snapshot, mc_snapshot
    return _workbook_bytes(
        build_term_workbook_from_spec(
            spec,
            out_path=out_path,
            alm_snapshot=alm_snapshot,
            alm_assumptions=alm_assumptions,
        )
    )


@register_builder(ProductType.RILA, spec_type=RILAExcelBuildSpec)
def _build_rila_workbook(
    *,
    spec: Any,
    out_path: str | Path | None,
    python_snapshot: ExcelPythonSnapshot | None,
    mc_snapshot: MCExcelSnapshot | None,
    alm_snapshot: ALMExcelSnapshot | None,
    alm_assumptions: sp.ALMAssumptions | None,
) -> bytes:
    # RILA builder takes alm_* but not python_snapshot / mc_snapshot at
    # this layer (the RILA workbook embeds its own MC sheet from the
    # spec). Accept and drop them so the dispatcher signature is uniform.
    del python_snapshot, mc_snapshot
    return _workbook_bytes(
        build_rila_workbook_from_spec(
            spec,
            out_path=out_path,
            alm_snapshot=alm_snapshot,
            alm_assumptions=alm_assumptions,
        )
    )


# ---------------------------------------------------------------------------
# Seven new product builders. ProductType: MYGA / FIA / VARIABLE_ANNUITY /
# WHOLE_LIFE / UNIVERSAL_LIFE / INDEXED_UL / VARIABLE_UL.
# Each accepts the union dispatcher kwargs and forwards only what its own
# builder needs. None of the new builders embed ALM at this layer (v1).
# ---------------------------------------------------------------------------


@register_builder(ProductType.MYGA, spec_type=MYGAExcelBuildSpec)
def _build_myga_workbook(
    *,
    spec: Any,
    out_path: str | Path | None,
    python_snapshot: ExcelPythonSnapshot | None,
    mc_snapshot: MCExcelSnapshot | None,
    alm_snapshot: ALMExcelSnapshot | None,
    alm_assumptions: sp.ALMAssumptions | None,
) -> bytes:
    del python_snapshot, mc_snapshot, alm_snapshot, alm_assumptions
    return _workbook_bytes(build_myga_workbook_from_spec(spec, out_path=out_path))


@register_builder(ProductType.FIA, spec_type=FIAExcelBuildSpec)
def _build_fia_workbook(
    *,
    spec: Any,
    out_path: str | Path | None,
    python_snapshot: ExcelPythonSnapshot | None,
    mc_snapshot: MCExcelSnapshot | None,
    alm_snapshot: ALMExcelSnapshot | None,
    alm_assumptions: sp.ALMAssumptions | None,
) -> bytes:
    del python_snapshot, mc_snapshot, alm_snapshot, alm_assumptions
    return _workbook_bytes(build_fia_workbook_from_spec(spec, out_path=out_path))


@register_builder(ProductType.VARIABLE_ANNUITY, spec_type=VAExcelBuildSpec)
def _build_va_workbook(
    *,
    spec: Any,
    out_path: str | Path | None,
    python_snapshot: ExcelPythonSnapshot | None,
    mc_snapshot: MCExcelSnapshot | None,
    alm_snapshot: ALMExcelSnapshot | None,
    alm_assumptions: sp.ALMAssumptions | None,
) -> bytes:
    del python_snapshot, mc_snapshot, alm_snapshot, alm_assumptions
    return _workbook_bytes(build_va_workbook_from_spec(spec, out_path=out_path))


@register_builder(ProductType.WHOLE_LIFE, spec_type=WLExcelBuildSpec)
def _build_wl_workbook(
    *,
    spec: Any,
    out_path: str | Path | None,
    python_snapshot: ExcelPythonSnapshot | None,
    mc_snapshot: MCExcelSnapshot | None,
    alm_snapshot: ALMExcelSnapshot | None,
    alm_assumptions: sp.ALMAssumptions | None,
) -> bytes:
    del python_snapshot, mc_snapshot, alm_snapshot, alm_assumptions
    return _workbook_bytes(build_wl_workbook_from_spec(spec, out_path=out_path))


@register_builder(ProductType.UNIVERSAL_LIFE, spec_type=ULExcelBuildSpec)
def _build_ul_workbook(
    *,
    spec: Any,
    out_path: str | Path | None,
    python_snapshot: ExcelPythonSnapshot | None,
    mc_snapshot: MCExcelSnapshot | None,
    alm_snapshot: ALMExcelSnapshot | None,
    alm_assumptions: sp.ALMAssumptions | None,
) -> bytes:
    del python_snapshot, mc_snapshot, alm_snapshot, alm_assumptions
    return _workbook_bytes(build_ul_workbook_from_spec(spec, out_path=out_path))


@register_builder(ProductType.INDEXED_UL, spec_type=IULExcelBuildSpec)
def _build_iul_workbook(
    *,
    spec: Any,
    out_path: str | Path | None,
    python_snapshot: ExcelPythonSnapshot | None,
    mc_snapshot: MCExcelSnapshot | None,
    alm_snapshot: ALMExcelSnapshot | None,
    alm_assumptions: sp.ALMAssumptions | None,
) -> bytes:
    del python_snapshot, mc_snapshot, alm_snapshot, alm_assumptions
    return _workbook_bytes(build_iul_workbook_from_spec(spec, out_path=out_path))


@register_builder(ProductType.VARIABLE_UL, spec_type=VULExcelBuildSpec)
def _build_vul_workbook(
    *,
    spec: Any,
    out_path: str | Path | None,
    python_snapshot: ExcelPythonSnapshot | None,
    mc_snapshot: MCExcelSnapshot | None,
    alm_snapshot: ALMExcelSnapshot | None,
    alm_assumptions: sp.ALMAssumptions | None,
) -> bytes:
    del python_snapshot, mc_snapshot, alm_snapshot, alm_assumptions
    return _workbook_bytes(build_vul_workbook_from_spec(spec, out_path=out_path))


def registered_builders() -> tuple[ProductType, ...]:
    """Return the tuple of product types that have a registered builder.

    Public registration status is derived from ``ProductDefinition``. The
    private ``_BUILDER_REGISTRY`` remains as the decorator seed consumed by
    product shims during import.
    """
    from annuity_model.products import registered_product_types

    return tuple(registered_product_types())


def workbook_builders_by_type() -> dict[ProductType, ProductWorkbookBuilder]:
    """Compatibility copy derived from canonical ``ProductDefinition``."""
    from annuity_model.products import workbook_builders_by_type as canonical_view

    return dict(canonical_view())


def workbook_builder_spec_types_by_type() -> dict[ProductType, type | None]:
    """Compatibility copy derived from canonical ``ProductDefinition``."""
    from annuity_model.products import workbook_builder_spec_types_by_type as canonical_view

    return dict(canonical_view())


def workbook_builder_registry_view() -> MappingProxyType[ProductType, ProductWorkbookBuilder]:
    """Immutable builder compatibility view derived from ProductDefinition."""
    return MappingProxyType(workbook_builders_by_type())


def workbook_builder_spec_type_view() -> MappingProxyType[ProductType, type | None]:
    """Immutable spec-type compatibility view derived from ProductDefinition."""
    return MappingProxyType(workbook_builder_spec_types_by_type())


def build_product_workbook(
    *,
    product_type: ProductType,
    spec: Any,
    out_path: str | Path | None = None,
    python_snapshot: ExcelPythonSnapshot | None = None,
    mc_snapshot: MCExcelSnapshot | None = None,
    alm_snapshot: ALMExcelSnapshot | None = None,
    alm_assumptions: sp.ALMAssumptions | None = None,
) -> bytes:
    from annuity_model.products import get_product_definition

    try:
        definition = get_product_definition(product_type)
    except KeyError as exc:
        raise NotImplementedError(
            f"Workbook builder is not implemented for product '{product_type.value}'."
        ) from exc

    builder = definition.builder
    expected_spec = definition.builder_spec_type
    if expected_spec is not None and not isinstance(spec, expected_spec):
        raise TypeError(
            f"{product_type.value} workbook builder requires {expected_spec.__name__}, "
            f"got {type(spec).__name__}."
        )
    return builder(
        spec=spec,
        out_path=out_path,
        python_snapshot=python_snapshot,
        mc_snapshot=mc_snapshot,
        alm_snapshot=alm_snapshot,
        alm_assumptions=alm_assumptions,
    )
