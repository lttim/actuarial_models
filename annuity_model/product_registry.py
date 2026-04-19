from __future__ import annotations

from collections.abc import Callable, Mapping
from dataclasses import dataclass
from enum import Enum
from typing import Any, Protocol, Union

import numpy as np

import pricing_projection as sp
import rila_projection as rp
import term_projection as tp
from build_pricing_excel_workbook import ExcelBuildSpec, excel_spec_from_launcher
from build_rila_excel_workbook import RILAExcelBuildSpec, rila_excel_spec_from_launcher
from build_term_excel_workbook import TermExcelBuildSpec, term_excel_spec_from_launcher

# Union of every contract dataclass currently understood by an adapter.
# Tightening the ``ProductAdapter`` Protocol from ``contract: object`` to
# this union (P1, 2026-04) lets mypy catch "wrong product, wrong contract"
# wiring at the call site instead of at the runtime ``isinstance`` check
# inside each adapter. New products MUST extend this union when they land.
ProductContract = Union[sp.SPIAContract, tp.TermLifeContract, rp.RILAContract]


class ProductType(str, Enum):
    SPIA = "spia"
    TERM_LIFE = "term_life"
    RILA = "rila"
    WHOLE_LIFE = "whole_life"
    VARIABLE_ANNUITY = "variable_annuity"


@dataclass(frozen=True)
class ProductCapabilities:
    supports_economic_scenario: bool
    supports_monte_carlo: bool


@dataclass(frozen=True)
class TermContractUIConfig:
    death_benefit_label: str
    default_death_benefit: float
    term_length_options: tuple[str, ...]
    premium_mode_options: tuple[str, ...]
    benefit_timing_options: tuple[str, ...]
    default_monthly_premium: float


@dataclass(frozen=True)
class PricingMetric:
    label: str
    value: float
    is_money: bool


@dataclass(frozen=True)
class ProductUIConfig:
    selected_info_message: str | None
    projection_csv_filename: str
    recalc_workbook_filename: str


class ProductAdapter(Protocol):
    """Per-product pricing adapter Protocol.

    Tightened in P1 (2026-04):

    * ``contract`` is now :data:`ProductContract` (a union of every known
      contract dataclass) instead of ``object``. mypy catches mis-wired
      contracts at the call site; the existing runtime ``isinstance``
      checks inside each adapter implementation stay as a defense in
      depth.
    * Added :meth:`validate_run_inputs`, an optional pre-flight that the
      Streamlit UI invokes before constructing a contract from the
      session-state dict. Adapters that have no constraints can leave
      the default no-op (the Protocol satisfaction does not require
      overriding it).

    The result types stay typed as ``object`` at the Protocol level only
    because the four concrete return types (SPIA / Term / RILA pricing
    + monte carlo + spec) form a heterogeneous matrix; downstream code
    narrows by ``isinstance`` against the concrete result class.
    """

    @property
    def product_type(self) -> ProductType: ...

    @property
    def display_name(self) -> str: ...

    def is_available(self) -> bool: ...

    def price(
        self,
        *,
        contract: ProductContract,
        yield_curve: sp.YieldCurve,
        mortality: sp.MortalityTableQx | sp.MortalityTableRP2014MP2016,
        horizon_age: int,
        spread: float,
        valuation_year: int | None,
        expenses: sp.ExpenseAssumptions | None,
        expenses_csv_path: str,
        index_scenario_csv_path: str | None,
        expense_annual_inflation: float,
    ) -> object: ...

    def price_monte_carlo(
        self,
        *,
        contract: ProductContract,
        yield_curve: sp.YieldCurve,
        mortality: sp.MortalityTableQx | sp.MortalityTableRP2014MP2016,
        horizon_age: int,
        spread: float,
        valuation_year: int | None,
        expenses: sp.ExpenseAssumptions | None,
        expenses_csv_path: str,
        expense_annual_inflation: float,
        n_sims: int,
        annual_drift: float,
        annual_vol: float,
        seed: int,
        s0: float,
    ) -> object: ...

    def excel_spec_from_run(
        self,
        *,
        contract: ProductContract,
        yield_curve: sp.YieldCurve,
        mortality: sp.MortalityTableQx | sp.MortalityTableRP2014MP2016,
        horizon_age: int,
        spread: float,
        valuation_year: int,
        expenses: sp.ExpenseAssumptions,
        yield_mode_label: str,
        mortality_mode_label: str,
        expense_mode_label: str,
        index_s0: float,
        index_levels_at_payment: np.ndarray,
        expense_annual_inflation: float,
    ) -> ExcelBuildSpec | TermExcelBuildSpec | RILAExcelBuildSpec: ...


def validate_run_inputs(
    product_type: ProductType, state: Mapping[str, Any]
) -> tuple[str, ...]:
    """Per-product pre-flight for Streamlit run-form session state.

    Returns a tuple of human-readable error messages; empty tuple means
    the inputs are well-formed and the run may proceed. The Streamlit
    UI may surface the messages via ``st.error`` and refuse to construct
    the contract dataclass.

    P1 (2026-04) introduces this hook as a *single source of truth* for
    "is this run launchable" so the validation does not get scattered
    across pricing_ui.py with stringly-typed conditionals. Today the
    function applies only the cross-product invariants below; per-product
    rules (e.g. RILA cap > floor) can be folded in via the
    ``_PRODUCT_VALIDATORS`` registry.

    Cross-product invariants (always enforced):

    * ``issue_age`` and ``horizon_age`` must both be present and integer.
    * ``horizon_age > issue_age`` (otherwise the projection is empty).

    Per-product extensions land in :data:`_PRODUCT_VALIDATORS`.
    """
    errors: list[str] = []

    issue_age = state.get("issue_age")
    horizon_age = state.get("horizon_age")
    if issue_age is None:
        errors.append("issue_age is required.")
    if horizon_age is None:
        errors.append("horizon_age is required.")
    if isinstance(issue_age, (int, float)) and isinstance(horizon_age, (int, float)):
        if int(horizon_age) <= int(issue_age):
            errors.append(
                f"horizon_age ({int(horizon_age)}) must be strictly greater than "
                f"issue_age ({int(issue_age)}); otherwise the projection horizon is empty."
            )

    extra = _PRODUCT_VALIDATORS.get(product_type)
    if extra is not None:
        errors.extend(extra(state))

    return tuple(errors)


# Per-product validator registry. A new product MAY add an entry that
# returns the list of additional error messages (after the cross-product
# checks have run). Keeping it as a dict-of-callables keeps the dispatch
# branch-free; tests/test_meta_invariants.py asserts every implemented
# product is either present or explicitly absent.
_PRODUCT_VALIDATORS: dict[ProductType, Callable[[Mapping[str, Any]], list[str]]] = {}


@dataclass(frozen=True)
class SPIAProductAdapter:
    @property
    def product_type(self) -> ProductType:
        return ProductType.SPIA

    @property
    def display_name(self) -> str:
        return "SPIA"

    def is_available(self) -> bool:
        return True

    def price(
        self,
        *,
        contract: object,
        yield_curve: sp.YieldCurve,
        mortality: sp.MortalityTableQx | sp.MortalityTableRP2014MP2016,
        horizon_age: int,
        spread: float,
        valuation_year: int | None,
        expenses: sp.ExpenseAssumptions | None,
        expenses_csv_path: str,
        index_scenario_csv_path: str | None,
        expense_annual_inflation: float,
    ) -> sp.SPIAProjectionResult:
        if not isinstance(contract, sp.SPIAContract):
            raise TypeError("SPIA adapter requires SPIAContract.")
        return sp.price_spia_single_premium(
            contract=contract,
            yield_curve=yield_curve,
            mortality=mortality,
            horizon_age=horizon_age,
            spread=spread,
            valuation_year=valuation_year,
            expenses=expenses,
            expenses_csv_path=expenses_csv_path,
            index_scenario_csv_path=index_scenario_csv_path,
            expense_annual_inflation=expense_annual_inflation,
        )

    def price_monte_carlo(
        self,
        *,
        contract: object,
        yield_curve: sp.YieldCurve,
        mortality: sp.MortalityTableQx | sp.MortalityTableRP2014MP2016,
        horizon_age: int,
        spread: float,
        valuation_year: int | None,
        expenses: sp.ExpenseAssumptions | None,
        expenses_csv_path: str,
        expense_annual_inflation: float,
        n_sims: int,
        annual_drift: float,
        annual_vol: float,
        seed: int,
        s0: float,
    ) -> sp.SPIAMonteCarloResult:
        if not isinstance(contract, sp.SPIAContract):
            raise TypeError("SPIA adapter requires SPIAContract.")
        return sp.price_spia_single_premium_monte_carlo(
            contract=contract,
            yield_curve=yield_curve,
            mortality=mortality,
            horizon_age=horizon_age,
            spread=spread,
            valuation_year=valuation_year,
            expenses=expenses,
            expenses_csv_path=expenses_csv_path,
            expense_annual_inflation=expense_annual_inflation,
            n_sims=n_sims,
            annual_drift=annual_drift,
            annual_vol=annual_vol,
            seed=seed,
            s0=s0,
        )

    def excel_spec_from_run(
        self,
        *,
        contract: object,
        yield_curve: sp.YieldCurve,
        mortality: sp.MortalityTableQx | sp.MortalityTableRP2014MP2016,
        horizon_age: int,
        spread: float,
        valuation_year: int,
        expenses: sp.ExpenseAssumptions,
        yield_mode_label: str,
        mortality_mode_label: str,
        expense_mode_label: str,
        index_s0: float,
        index_levels_at_payment: np.ndarray,
        expense_annual_inflation: float,
    ) -> ExcelBuildSpec:
        if not isinstance(contract, sp.SPIAContract):
            raise TypeError("SPIA adapter requires SPIAContract.")
        return excel_spec_from_launcher(
            contract=contract,
            yield_curve=yield_curve,
            mortality=mortality,
            horizon_age=horizon_age,
            spread=spread,
            valuation_year=valuation_year,
            expenses=expenses,
            yield_mode_label=yield_mode_label,
            mortality_mode_label=mortality_mode_label,
            expense_mode_label=expense_mode_label,
            index_s0=index_s0,
            index_levels_at_payment=index_levels_at_payment,
            expense_annual_inflation=expense_annual_inflation,
        )


_SPIA_ADAPTER = SPIAProductAdapter()


@dataclass(frozen=True)
class TermLifeProductAdapter:
    @property
    def product_type(self) -> ProductType:
        return ProductType.TERM_LIFE

    @property
    def display_name(self) -> str:
        return "Term Life (20Y)"

    def is_available(self) -> bool:
        return True

    def price(
        self,
        *,
        contract: object,
        yield_curve: sp.YieldCurve,
        mortality: sp.MortalityTableQx | sp.MortalityTableRP2014MP2016,
        horizon_age: int,
        spread: float,
        valuation_year: int | None,
        expenses: sp.ExpenseAssumptions | None,
        expenses_csv_path: str,
        index_scenario_csv_path: str | None,
        expense_annual_inflation: float,
    ) -> tp.TermLifeProjectionResult:
        del expenses, expenses_csv_path, index_scenario_csv_path, expense_annual_inflation
        if not isinstance(contract, tp.TermLifeContract):
            raise TypeError("Term adapter requires TermLifeContract.")
        return tp.price_term_life_level_monthly(
            contract=contract,
            yield_curve=yield_curve,
            mortality=mortality,
            horizon_age=horizon_age,
            spread=spread,
            valuation_year=valuation_year,
        )

    def price_monte_carlo(
        self,
        *,
        contract: object,
        yield_curve: sp.YieldCurve,
        mortality: sp.MortalityTableQx | sp.MortalityTableRP2014MP2016,
        horizon_age: int,
        spread: float,
        valuation_year: int | None,
        expenses: sp.ExpenseAssumptions | None,
        expenses_csv_path: str,
        expense_annual_inflation: float,
        n_sims: int,
        annual_drift: float,
        annual_vol: float,
        seed: int,
        s0: float,
    ) -> object:
        del (
            contract,
            yield_curve,
            mortality,
            horizon_age,
            spread,
            valuation_year,
            expenses,
            expenses_csv_path,
            expense_annual_inflation,
            n_sims,
            annual_drift,
            annual_vol,
            seed,
            s0,
        )
        raise NotImplementedError("Monte Carlo is not implemented for Term Life in this release.")

    def excel_spec_from_run(
        self,
        *,
        contract: object,
        yield_curve: sp.YieldCurve,
        mortality: sp.MortalityTableQx | sp.MortalityTableRP2014MP2016,
        horizon_age: int,
        spread: float,
        valuation_year: int,
        expenses: sp.ExpenseAssumptions,
        yield_mode_label: str,
        mortality_mode_label: str,
        expense_mode_label: str,
        index_s0: float,
        index_levels_at_payment: np.ndarray,
        expense_annual_inflation: float,
    ) -> TermExcelBuildSpec:
        del index_s0, index_levels_at_payment, expense_annual_inflation
        if not isinstance(contract, tp.TermLifeContract):
            raise TypeError("Term adapter requires TermLifeContract.")
        return term_excel_spec_from_launcher(
            contract=contract,
            yield_curve=yield_curve,
            mortality=mortality,
            horizon_age=horizon_age,
            spread=spread,
            valuation_year=valuation_year,
            expenses=expenses,
            yield_mode_label=yield_mode_label,
            mortality_mode_label=mortality_mode_label,
            expense_mode_label=expense_mode_label,
        )


_TERM_ADAPTER = TermLifeProductAdapter()


@dataclass(frozen=True)
class RILAProductAdapter:
    @property
    def product_type(self) -> ProductType:
        return ProductType.RILA

    @property
    def display_name(self) -> str:
        return "RILA (accumulation)"

    def is_available(self) -> bool:
        return True

    def price(
        self,
        *,
        contract: object,
        yield_curve: sp.YieldCurve,
        mortality: sp.MortalityTableQx | sp.MortalityTableRP2014MP2016,
        horizon_age: int,
        spread: float,
        valuation_year: int | None,
        expenses: sp.ExpenseAssumptions | None,
        expenses_csv_path: str,
        index_scenario_csv_path: str | None,
        expense_annual_inflation: float,
    ) -> rp.RILAProjectionResult:
        if not isinstance(contract, rp.RILAContract):
            raise TypeError("RILA adapter requires RILAContract.")
        return rp.price_rila_single_premium(
            contract=contract,
            yield_curve=yield_curve,
            mortality=mortality,
            horizon_age=horizon_age,
            spread=spread,
            valuation_year=valuation_year,
            expenses=expenses,
            expenses_csv_path=expenses_csv_path,
            index_scenario_csv_path=index_scenario_csv_path,
            expense_annual_inflation=expense_annual_inflation,
        )

    def price_monte_carlo(
        self,
        *,
        contract: object,
        yield_curve: sp.YieldCurve,
        mortality: sp.MortalityTableQx | sp.MortalityTableRP2014MP2016,
        horizon_age: int,
        spread: float,
        valuation_year: int | None,
        expenses: sp.ExpenseAssumptions | None,
        expenses_csv_path: str,
        expense_annual_inflation: float,
        n_sims: int,
        annual_drift: float,
        annual_vol: float,
        seed: int,
        s0: float,
    ) -> rp.RILAMonteCarloResult:
        if not isinstance(contract, rp.RILAContract):
            raise TypeError("RILA adapter requires RILAContract.")
        return rp.price_rila_single_premium_monte_carlo(
            contract=contract,
            yield_curve=yield_curve,
            mortality=mortality,
            horizon_age=horizon_age,
            spread=spread,
            valuation_year=valuation_year,
            expenses=expenses,
            expenses_csv_path=expenses_csv_path,
            expense_annual_inflation=expense_annual_inflation,
            n_sims=n_sims,
            annual_drift=annual_drift,
            annual_vol=annual_vol,
            seed=seed,
            s0=s0,
        )

    def excel_spec_from_run(
        self,
        *,
        contract: object,
        yield_curve: sp.YieldCurve,
        mortality: sp.MortalityTableQx | sp.MortalityTableRP2014MP2016,
        horizon_age: int,
        spread: float,
        valuation_year: int,
        expenses: sp.ExpenseAssumptions,
        yield_mode_label: str,
        mortality_mode_label: str,
        expense_mode_label: str,
        index_s0: float,
        index_levels_at_payment: np.ndarray,
        expense_annual_inflation: float,
    ) -> RILAExcelBuildSpec:
        if not isinstance(contract, rp.RILAContract):
            raise TypeError("RILA adapter requires RILAContract.")
        return rila_excel_spec_from_launcher(
            contract=contract,
            yield_curve=yield_curve,
            mortality=mortality,
            horizon_age=horizon_age,
            spread=spread,
            valuation_year=valuation_year,
            expenses=expenses,
            yield_mode_label=yield_mode_label,
            mortality_mode_label=mortality_mode_label,
            expense_mode_label=expense_mode_label,
            index_s0=index_s0,
            index_levels_at_payment=index_levels_at_payment,
            expense_annual_inflation=expense_annual_inflation,
        )


_RILA_ADAPTER = RILAProductAdapter()

_PRODUCT_DISPLAY_NAME: dict[ProductType, str] = {
    ProductType.SPIA: "SPIA",
    ProductType.TERM_LIFE: "Term Life (20Y)",
    ProductType.RILA: "RILA (accumulation)",
    ProductType.WHOLE_LIFE: "Whole Life (coming soon)",
    ProductType.VARIABLE_ANNUITY: "Variable Annuity (coming soon)",
}

_PRODUCT_CAPABILITIES: dict[ProductType, ProductCapabilities] = {
    ProductType.SPIA: ProductCapabilities(
        supports_economic_scenario=True, supports_monte_carlo=True
    ),
    ProductType.TERM_LIFE: ProductCapabilities(
        supports_economic_scenario=False, supports_monte_carlo=False
    ),
    ProductType.RILA: ProductCapabilities(
        supports_economic_scenario=True, supports_monte_carlo=True
    ),
    ProductType.WHOLE_LIFE: ProductCapabilities(
        supports_economic_scenario=False, supports_monte_carlo=False
    ),
    ProductType.VARIABLE_ANNUITY: ProductCapabilities(
        supports_economic_scenario=True, supports_monte_carlo=True
    ),
}

_PRODUCT_MORTALITY_MODE_OPTIONS: dict[ProductType, tuple[str, ...]] = {
    ProductType.SPIA: ("synthetic", "qx_csv", "rp2014_mp2016"),
    ProductType.TERM_LIFE: ("us_ssa_2015_period", "qx_csv", "synthetic"),
    ProductType.RILA: ("synthetic", "qx_csv", "rp2014_mp2016"),
    ProductType.WHOLE_LIFE: ("synthetic", "qx_csv"),
    ProductType.VARIABLE_ANNUITY: ("synthetic", "qx_csv", "rp2014_mp2016"),
}

_PRODUCT_DEFAULT_MORTALITY_MODE: dict[ProductType, str] = {
    ProductType.SPIA: "rp2014_mp2016",
    ProductType.TERM_LIFE: "us_ssa_2015_period",
    ProductType.RILA: "rp2014_mp2016",
    ProductType.WHOLE_LIFE: "synthetic",
    ProductType.VARIABLE_ANNUITY: "rp2014_mp2016",
}

_MORTALITY_MODE_LABELS: dict[str, str] = {
    "synthetic": "Synthetic (demo, wide age range)",
    "qx_csv": "Static q_x CSV",
    "rp2014_mp2016": "RP-2014 Healthy Male + MP-2016 (xlsx or cached CSV)",
    "us_ssa_2015_period": "US SSA 2015 period life table (sex-specific default for Term)",
}

_TERM_CONTRACT_UI_CONFIG = TermContractUIConfig(
    death_benefit_label="Death benefit ($)",
    default_death_benefit=250_000.0,
    # NOTE: option tuples are kept in sync with the parser maps below by
    # ``test_pricing_ui_term_config.test_term_ui_config_options_match_parser_maps``.
    term_length_options=("20 years",),
    premium_mode_options=("Level monthly",),
    benefit_timing_options=("EOY death benefit",),
    default_monthly_premium=250.0,
)


# UI-label → engine-value mappings for the Term contract.
#
# The Streamlit selectboxes display human-readable labels (e.g. "20 years"),
# but the ``TermLifeContract`` engine dataclass takes typed scalars
# (``term_years: int``, ``premium_mode: Literal["level_monthly"]``,
# ``benefit_timing: Literal["eoy_death"]``). These maps are the *single source
# of truth* for that translation; ``pricing_ui.py`` MUST round-trip every
# widget value through them rather than hard-coding ``term_years=20`` etc.
#
# When new options are added to ``TermContractUIConfig.*_options``, add the
# matching label here in the same PR. ``test_pricing_ui_term_config.py``
# enforces that every option label is parseable.
_TERM_LENGTH_YEARS_BY_LABEL: dict[str, int] = {
    "20 years": 20,
}
_TERM_PREMIUM_MODE_BY_LABEL: dict[str, str] = {
    "Level monthly": "level_monthly",
}
_TERM_BENEFIT_TIMING_BY_LABEL: dict[str, str] = {
    "EOY death benefit": "eoy_death",
}


def parse_term_length_label_to_years(label: str) -> int:
    try:
        return _TERM_LENGTH_YEARS_BY_LABEL[label]
    except KeyError as exc:
        raise ValueError(
            f"Unknown Term length label {label!r}; "
            f"add a mapping in product_registry._TERM_LENGTH_YEARS_BY_LABEL."
        ) from exc


def parse_term_premium_mode_label(label: str) -> str:
    try:
        return _TERM_PREMIUM_MODE_BY_LABEL[label]
    except KeyError as exc:
        raise ValueError(
            f"Unknown Term premium mode label {label!r}; "
            f"add a mapping in product_registry._TERM_PREMIUM_MODE_BY_LABEL."
        ) from exc


def parse_term_benefit_timing_label(label: str) -> str:
    try:
        return _TERM_BENEFIT_TIMING_BY_LABEL[label]
    except KeyError as exc:
        raise ValueError(
            f"Unknown Term benefit timing label {label!r}; "
            f"add a mapping in product_registry._TERM_BENEFIT_TIMING_BY_LABEL."
        ) from exc


def term_length_label_options() -> tuple[str, ...]:
    return tuple(_TERM_LENGTH_YEARS_BY_LABEL.keys())


def term_premium_mode_label_options() -> tuple[str, ...]:
    return tuple(_TERM_PREMIUM_MODE_BY_LABEL.keys())


def term_benefit_timing_label_options() -> tuple[str, ...]:
    return tuple(_TERM_BENEFIT_TIMING_BY_LABEL.keys())

_PRODUCT_UI_CONFIG: dict[ProductType, ProductUIConfig] = {
    ProductType.SPIA: ProductUIConfig(
        selected_info_message=None,
        projection_csv_filename="pricing_projection_spia.csv",
        recalc_workbook_filename="spia_recalc_model.xlsx",
    ),
    ProductType.TERM_LIFE: ProductUIConfig(
        selected_info_message="Term Life (20Y) is enabled with deterministic pricing. Monte Carlo is not available in this release.",
        projection_csv_filename="pricing_projection_term_life.csv",
        recalc_workbook_filename="term_life_recalc_model.xlsx",
    ),
    ProductType.RILA: ProductUIConfig(
        selected_info_message=None,
        projection_csv_filename="pricing_projection_rila.csv",
        recalc_workbook_filename="rila_recalc_model.xlsx",
    ),
    ProductType.WHOLE_LIFE: ProductUIConfig(
        selected_info_message="Selected product is scaffolded but not implemented yet.",
        projection_csv_filename="pricing_projection_whole_life.csv",
        recalc_workbook_filename="whole_life_recalc_model.xlsx",
    ),
    ProductType.VARIABLE_ANNUITY: ProductUIConfig(
        selected_info_message="Selected product is scaffolded but not implemented yet.",
        projection_csv_filename="pricing_projection_variable_annuity.csv",
        recalc_workbook_filename="variable_annuity_recalc_model.xlsx",
    ),
}


# Pluggable adapter registry. Each implemented product registers its adapter
# instance here. Adding a new product means appending one entry; the dispatcher
# below stays product-agnostic. (Future P3: replace with a @register decorator
# pattern + auto-discovery once we move to the src/ layout.)
_PRODUCT_ADAPTERS: dict[ProductType, ProductAdapter] = {
    ProductType.SPIA: _SPIA_ADAPTER,
    ProductType.TERM_LIFE: _TERM_ADAPTER,
    ProductType.RILA: _RILA_ADAPTER,
}


def get_product_adapter(product_type: ProductType) -> ProductAdapter:
    """Return the adapter for *product_type* or raise NotImplementedError.

    Unimplemented products (Whole Life, Variable Annuity) are present in
    :class:`ProductType` so the UI can render disabled options, but they
    have no entry in :data:`_PRODUCT_ADAPTERS`.
    """
    adapter = _PRODUCT_ADAPTERS.get(product_type)
    if adapter is None:
        raise NotImplementedError(f"{_PRODUCT_DISPLAY_NAME[product_type]} is not implemented yet.")
    return adapter


def implemented_product_types() -> tuple[ProductType, ...]:
    """Return the tuple of product types with a registered adapter.

    Used by meta-tests in Phase 4 to assert every implemented product also
    has an entry in ``LIABILITY_LAYOUTS``.
    """
    return tuple(_PRODUCT_ADAPTERS)


def product_options_for_ui() -> list[ProductType]:
    return [
        ProductType.SPIA,
        ProductType.TERM_LIFE,
        ProductType.RILA,
        ProductType.WHOLE_LIFE,
        ProductType.VARIABLE_ANNUITY,
    ]


def product_label(product_type: ProductType) -> str:
    return _PRODUCT_DISPLAY_NAME[product_type]


def get_product_capabilities(product_type: ProductType) -> ProductCapabilities:
    return _PRODUCT_CAPABILITIES[product_type]


def get_product_mortality_mode_options(product_type: ProductType) -> tuple[str, ...]:
    return _PRODUCT_MORTALITY_MODE_OPTIONS[product_type]


def get_product_default_mortality_mode(product_type: ProductType) -> str:
    return _PRODUCT_DEFAULT_MORTALITY_MODE[product_type]


def get_mortality_mode_label(mode: str) -> str:
    return _MORTALITY_MODE_LABELS.get(mode, mode)


def get_term_contract_ui_config() -> TermContractUIConfig:
    return _TERM_CONTRACT_UI_CONFIG


def _spia_pricing_metrics(result: Any) -> tuple[PricingMetric, ...]:
    return (
        PricingMetric(label="Single premium", value=float(result.single_premium), is_money=True),
        PricingMetric(label="PV benefit", value=float(result.pv_benefit), is_money=True),
        PricingMetric(
            label="PV monthly expenses", value=float(result.pv_monthly_expenses), is_money=True
        ),
        PricingMetric(label="Annuity factor", value=float(result.annuity_factor), is_money=False),
    )


def _rila_pricing_metrics(result: Any) -> tuple[PricingMetric, ...]:
    return (
        PricingMetric(label="Single premium", value=float(result.single_premium), is_money=True),
        PricingMetric(label="PV benefit (claims)", value=float(result.pv_benefit), is_money=True),
        PricingMetric(
            label="PV monthly expenses", value=float(result.pv_monthly_expenses), is_money=True
        ),
        PricingMetric(label="Annuity factor", value=float(result.annuity_factor), is_money=False),
    )


def _term_pricing_metrics(result: Any) -> tuple[PricingMetric, ...]:
    pv_claims = float(result.pv_benefit)
    pv_premiums = float(-float(result.pv_monthly_expenses))
    net_pv = float(result.single_premium)
    economic_reserve = np.asarray(getattr(result, "economic_reserve", np.asarray([], dtype=float)))
    issue_reserve = float(economic_reserve[0]) if economic_reserve.size else float("nan")
    return (
        PricingMetric(label="PV claims", value=pv_claims, is_money=True),
        PricingMetric(label="PV premiums", value=pv_premiums, is_money=True),
        PricingMetric(label="Net PV (claims - premiums)", value=net_pv, is_money=True),
        PricingMetric(label="Issue reserve", value=issue_reserve, is_money=True),
    )


# Per-product pricing-metric formatters. Adding a new product means one
# additional entry here; the dispatcher stays branch-free.
_PRICING_METRIC_FORMATTERS: dict[ProductType, Callable[[Any], tuple[PricingMetric, ...]]] = {
    ProductType.SPIA: _spia_pricing_metrics,
    ProductType.TERM_LIFE: _term_pricing_metrics,
    ProductType.RILA: _rila_pricing_metrics,
}


def get_pricing_metrics(product_type: ProductType, result: Any) -> tuple[PricingMetric, ...]:
    """Return the per-product pricing metric tuple.

    No silent fallback: if a product is implemented (in
    ``implemented_product_types()``) but missing from
    ``_PRICING_METRIC_FORMATTERS``, we raise a ``KeyError`` with a clear
    pointer to the registry instead of silently returning SPIA-shaped metrics
    that would mislead the UI. The completeness invariant is enforced by
    ``tests/test_metric_formatter_completeness.py``.
    """
    formatter = _PRICING_METRIC_FORMATTERS.get(product_type)
    if formatter is None:
        raise KeyError(
            f"No pricing-metric formatter registered for {product_type.value!r}; "
            f"add an entry in product_registry._PRICING_METRIC_FORMATTERS."
        )
    return formatter(result)


def get_product_ui_config(product_type: ProductType) -> ProductUIConfig:
    return _PRODUCT_UI_CONFIG[product_type]
