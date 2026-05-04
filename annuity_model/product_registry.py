from __future__ import annotations

from collections.abc import Callable, Mapping
from dataclasses import dataclass
from enum import Enum
from types import MappingProxyType
from typing import Any, Protocol, TypeAlias

import numpy as np

import fia_projection as fp
import iul_projection as iul
import myga_projection as my
import pricing_projection as sp
import rila_projection as rp
import term_projection as tp
import ul_projection as ul
import va_projection as va
import vul_projection as vul
import wl_projection as wl
from build_fia_excel_workbook import FIAExcelBuildSpec, fia_excel_spec_from_launcher
from build_iul_excel_workbook import IULExcelBuildSpec, iul_excel_spec_from_launcher
from build_myga_excel_workbook import MYGAExcelBuildSpec, myga_excel_spec_from_launcher
from build_pricing_excel_workbook import ExcelBuildSpec, excel_spec_from_launcher
from build_rila_excel_workbook import RILAExcelBuildSpec, rila_excel_spec_from_launcher
from build_term_excel_workbook import TermExcelBuildSpec, term_excel_spec_from_launcher
from build_ul_excel_workbook import ULExcelBuildSpec, ul_excel_spec_from_launcher
from build_va_excel_workbook import VAExcelBuildSpec, va_excel_spec_from_launcher
from build_vul_excel_workbook import VULExcelBuildSpec, vul_excel_spec_from_launcher
from build_wl_excel_workbook import WLExcelBuildSpec, wl_excel_spec_from_launcher

# Union of every contract dataclass currently understood by an adapter.
# Tightening the ``ProductAdapter`` Protocol from ``contract: object`` to
# this union (P1, 2026-04) lets mypy catch "wrong product, wrong contract"
# wiring at the call site instead of at the runtime ``isinstance`` check
# inside each adapter. New products MUST extend this union when they land.
ProductContract: TypeAlias = (
    sp.SPIAContract
    | tp.TermLifeContract
    | rp.RILAContract
    | my.MYGAContract
    | fp.FIAContract
    | va.VAContract
    | wl.WLContract
    | ul.ULContract
    | iul.IULContract
    | vul.VULContract
)


class ProductType(str, Enum):
    SPIA = "spia"
    TERM_LIFE = "term_life"
    RILA = "rila"
    WHOLE_LIFE = "whole_life"
    VARIABLE_ANNUITY = "variable_annuity"
    MYGA = "myga"
    FIA = "fia"
    UNIVERSAL_LIFE = "universal_life"
    INDEXED_UL = "indexed_ul"
    VARIABLE_UL = "variable_ul"


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


def validate_run_inputs(product_type: ProductType, state: Mapping[str, Any]) -> tuple[str, ...]:
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
    if (
        isinstance(issue_age, (int, float))
        and isinstance(horizon_age, (int, float))
        and int(horizon_age) <= int(issue_age)
    ):
        errors.append(
            f"horizon_age ({int(horizon_age)}) must be strictly greater than "
            f"issue_age ({int(issue_age)}); otherwise the projection horizon is empty."
        )

    validator = _PRODUCT_VALIDATORS.get(product_type)
    if validator is not None:
        extra_errors = validator(state)
        errors.extend(extra_errors)

    return tuple(errors)


def _validate_myga(state: Mapping[str, Any]) -> list[str]:
    errors: list[str] = []
    sp_val = state.get("myga_single_premium")
    rate = state.get("myga_declared_rate")
    years = state.get("myga_guarantee_years")
    if sp_val is not None and (not isinstance(sp_val, (int, float)) or float(sp_val) <= 0):
        errors.append("myga_single_premium must be > 0.")
    if rate is not None and (
        not isinstance(rate, (int, float)) or not (-0.5 <= float(rate) <= 1.0)
    ):
        errors.append("myga_declared_rate must be in [-0.5, 1.0].")
    if years is not None and (not isinstance(years, (int, float)) or not (1 <= int(years) <= 30)):
        errors.append("myga_guarantee_years must be in [1, 30].")
    return errors


def _validate_fia(state: Mapping[str, Any]) -> list[str]:
    errors: list[str] = []
    cap = state.get("fia_cap")
    floor = state.get("fia_floor")
    sp_val = state.get("fia_single_premium")
    if sp_val is not None and (not isinstance(sp_val, (int, float)) or float(sp_val) <= 0):
        errors.append("fia_single_premium must be > 0.")
    if cap is not None and floor is not None and float(cap) < float(floor):
        errors.append(f"fia cap ({cap}) must be >= floor ({floor}).")
    return errors


def _validate_va(state: Mapping[str, Any]) -> list[str]:
    errors: list[str] = []
    sp_val = state.get("va_single_premium")
    me = state.get("va_me_charge_annual")
    if sp_val is not None and (not isinstance(sp_val, (int, float)) or float(sp_val) <= 0):
        errors.append("va_single_premium must be > 0.")
    if me is not None and not (0.0 <= float(me) <= 0.05):
        errors.append("va_me_charge_annual must be in [0, 0.05] (5% cap).")
    return errors


def _validate_life_face_and_premium(state: Mapping[str, Any], prefix: str) -> list[str]:
    errors: list[str] = []
    face = state.get(f"{prefix}_face_amount")
    sp_val = state.get(f"{prefix}_single_premium")
    load = state.get(f"{prefix}_premium_load_pct")
    if face is not None and (not isinstance(face, (int, float)) or float(face) <= 0):
        errors.append(f"{prefix}_face_amount must be > 0.")
    if sp_val is not None and (not isinstance(sp_val, (int, float)) or float(sp_val) <= 0):
        errors.append(f"{prefix}_single_premium must be > 0.")
    if load is not None and not (0.0 <= float(load) < 1.0):
        errors.append(f"{prefix}_premium_load_pct must be in [0, 1).")
    return errors


def _validate_wl(state: Mapping[str, Any]) -> list[str]:
    errors: list[str] = []
    face = state.get("wl_face_amount")
    if face is not None and (not isinstance(face, (int, float)) or float(face) <= 0):
        errors.append("wl_face_amount must be > 0.")
    return errors


def _validate_ul(state: Mapping[str, Any]) -> list[str]:
    return _validate_life_face_and_premium(state, "ul")


def _validate_iul(state: Mapping[str, Any]) -> list[str]:
    errors = _validate_life_face_and_premium(state, "iul")
    cap = state.get("iul_cap")
    floor = state.get("iul_floor")
    if cap is not None and floor is not None and float(cap) < float(floor):
        errors.append(f"iul cap ({cap}) must be >= floor ({floor}).")
    return errors


def _validate_vul(state: Mapping[str, Any]) -> list[str]:
    return _validate_life_face_and_premium(state, "vul")


# Per-product validator registry. A new product MAY add an entry that
# returns the list of additional error messages (after the cross-product
# checks have run). Keeping it as a dict-of-callables keeps the dispatch
# branch-free; tests/test_meta_invariants.py asserts every implemented
# product is either present or explicitly absent.
_PRODUCT_VALIDATORS: dict[ProductType, Callable[[Mapping[str, Any]], list[str]]] = {
    ProductType.MYGA: _validate_myga,
    ProductType.FIA: _validate_fia,
    ProductType.VARIABLE_ANNUITY: _validate_va,
    ProductType.WHOLE_LIFE: _validate_wl,
    ProductType.UNIVERSAL_LIFE: _validate_ul,
    ProductType.INDEXED_UL: _validate_iul,
    ProductType.VARIABLE_UL: _validate_vul,
}


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


# ---------------------------------------------------------------------------
# Seven new product adapters (Phase 1-7 of seven_product_rollout_plan.md).
# Each adapter follows the SPIA/Term/RILA pattern: isinstance-check the
# contract, dispatch to the engine, accept the union of legacy adapter
# kwargs, and route Monte Carlo via the dedicated function (or raise
# NotImplementedError for products without MC).
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class _SimpleAccumulationAdapter:
    """Shared adapter for products whose engines take the same kwargs.

    All four accumulation products (MYGA / FIA / VA) AND the three life
    products (WL / UL / IUL / VUL) share the price() signature; only the
    contract type and engine function differ. Codifying the dispatch
    here cuts ~150 LOC of repetitive adapter boilerplate.

    Sub-classed below as ``_make_adapter(<ProductType>, <ContractCls>,
    <price_fn>, <mc_fn or None>, <spec_fn>, <SpecCls>)``.
    """

    product_type_value: ProductType
    display_name_value: str
    contract_type: type
    price_fn: Callable[..., Any]
    mc_fn: Callable[..., Any] | None
    spec_fn: Callable[..., Any]
    spec_type: type

    @property
    def product_type(self) -> ProductType:
        return self.product_type_value

    @property
    def display_name(self) -> str:
        return self.display_name_value

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
    ) -> object:
        if not isinstance(contract, self.contract_type):
            raise TypeError(
                f"{self.display_name_value} adapter requires "
                f"{self.contract_type.__name__}, got {type(contract).__name__}."
            )
        return self.price_fn(
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
    ) -> object:
        if self.mc_fn is None:
            raise NotImplementedError(
                f"{self.display_name_value} does not support Monte Carlo in this release."
            )
        if not isinstance(contract, self.contract_type):
            raise TypeError(
                f"{self.display_name_value} adapter requires "
                f"{self.contract_type.__name__}, got {type(contract).__name__}."
            )
        return self.mc_fn(
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
    ) -> Any:
        if not isinstance(contract, self.contract_type):
            raise TypeError(
                f"{self.display_name_value} adapter requires "
                f"{self.contract_type.__name__}, got {type(contract).__name__}."
            )
        return self.spec_fn(
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


_MYGA_ADAPTER = _SimpleAccumulationAdapter(
    product_type_value=ProductType.MYGA,
    display_name_value="MYGA",
    contract_type=my.MYGAContract,
    price_fn=my.price_myga_single_premium,
    mc_fn=None,
    spec_fn=myga_excel_spec_from_launcher,
    spec_type=MYGAExcelBuildSpec,
)
_FIA_ADAPTER = _SimpleAccumulationAdapter(
    product_type_value=ProductType.FIA,
    display_name_value="FIA",
    contract_type=fp.FIAContract,
    price_fn=fp.price_fia_single_premium,
    mc_fn=fp.price_fia_single_premium_monte_carlo,
    spec_fn=fia_excel_spec_from_launcher,
    spec_type=FIAExcelBuildSpec,
)
_VA_ADAPTER = _SimpleAccumulationAdapter(
    product_type_value=ProductType.VARIABLE_ANNUITY,
    display_name_value="Variable Annuity",
    contract_type=va.VAContract,
    price_fn=va.price_va_single_premium,
    mc_fn=va.price_va_single_premium_monte_carlo,
    spec_fn=va_excel_spec_from_launcher,
    spec_type=VAExcelBuildSpec,
)
_WL_ADAPTER = _SimpleAccumulationAdapter(
    product_type_value=ProductType.WHOLE_LIFE,
    display_name_value="Whole Life",
    contract_type=wl.WLContract,
    price_fn=wl.price_wl_single_premium,
    mc_fn=None,
    spec_fn=wl_excel_spec_from_launcher,
    spec_type=WLExcelBuildSpec,
)
_UL_ADAPTER = _SimpleAccumulationAdapter(
    product_type_value=ProductType.UNIVERSAL_LIFE,
    display_name_value="Universal Life",
    contract_type=ul.ULContract,
    price_fn=ul.price_ul_single_premium,
    mc_fn=None,
    spec_fn=ul_excel_spec_from_launcher,
    spec_type=ULExcelBuildSpec,
)
_IUL_ADAPTER = _SimpleAccumulationAdapter(
    product_type_value=ProductType.INDEXED_UL,
    display_name_value="Indexed UL",
    contract_type=iul.IULContract,
    price_fn=iul.price_iul_single_premium,
    mc_fn=iul.price_iul_single_premium_monte_carlo,
    spec_fn=iul_excel_spec_from_launcher,
    spec_type=IULExcelBuildSpec,
)
_VUL_ADAPTER = _SimpleAccumulationAdapter(
    product_type_value=ProductType.VARIABLE_UL,
    display_name_value="Variable UL",
    contract_type=vul.VULContract,
    price_fn=vul.price_vul_single_premium,
    mc_fn=vul.price_vul_single_premium_monte_carlo,
    spec_fn=vul_excel_spec_from_launcher,
    spec_type=VULExcelBuildSpec,
)

_PRODUCT_DISPLAY_NAME: dict[ProductType, str] = {
    ProductType.SPIA: "SPIA",
    ProductType.TERM_LIFE: "Term Life (20Y)",
    ProductType.RILA: "RILA (accumulation)",
    ProductType.WHOLE_LIFE: "Whole Life (single premium)",
    ProductType.VARIABLE_ANNUITY: "Variable Annuity (single premium)",
    ProductType.MYGA: "MYGA (multi-year guaranteed)",
    ProductType.FIA: "FIA (fixed indexed annuity)",
    ProductType.UNIVERSAL_LIFE: "Universal Life (single premium)",
    ProductType.INDEXED_UL: "Indexed UL (IUL, single premium)",
    ProductType.VARIABLE_UL: "Variable UL (VUL, single premium)",
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
    ProductType.MYGA: ProductCapabilities(
        supports_economic_scenario=False, supports_monte_carlo=False
    ),
    ProductType.FIA: ProductCapabilities(
        supports_economic_scenario=True, supports_monte_carlo=True
    ),
    ProductType.UNIVERSAL_LIFE: ProductCapabilities(
        supports_economic_scenario=False, supports_monte_carlo=False
    ),
    ProductType.INDEXED_UL: ProductCapabilities(
        supports_economic_scenario=True, supports_monte_carlo=True
    ),
    ProductType.VARIABLE_UL: ProductCapabilities(
        supports_economic_scenario=True, supports_monte_carlo=True
    ),
}

_PRODUCT_MORTALITY_MODE_OPTIONS: dict[ProductType, tuple[str, ...]] = {
    ProductType.SPIA: ("synthetic", "qx_csv", "rp2014_mp2016"),
    ProductType.TERM_LIFE: ("us_ssa_2015_period", "qx_csv", "synthetic"),
    ProductType.RILA: ("synthetic", "qx_csv", "rp2014_mp2016"),
    ProductType.WHOLE_LIFE: ("cso_2017_ult", "qx_csv", "synthetic"),
    ProductType.VARIABLE_ANNUITY: ("synthetic", "qx_csv", "rp2014_mp2016"),
    ProductType.MYGA: ("synthetic", "qx_csv", "rp2014_mp2016"),
    ProductType.FIA: ("synthetic", "qx_csv", "rp2014_mp2016"),
    ProductType.UNIVERSAL_LIFE: ("cso_2017_ult", "qx_csv", "synthetic"),
    ProductType.INDEXED_UL: ("cso_2017_ult", "qx_csv", "synthetic"),
    ProductType.VARIABLE_UL: ("cso_2017_ult", "qx_csv", "synthetic"),
}

_PRODUCT_DEFAULT_MORTALITY_MODE: dict[ProductType, str] = {
    ProductType.SPIA: "rp2014_mp2016",
    ProductType.TERM_LIFE: "us_ssa_2015_period",
    ProductType.RILA: "rp2014_mp2016",
    ProductType.WHOLE_LIFE: "cso_2017_ult",
    ProductType.VARIABLE_ANNUITY: "rp2014_mp2016",
    ProductType.MYGA: "rp2014_mp2016",
    ProductType.FIA: "rp2014_mp2016",
    ProductType.UNIVERSAL_LIFE: "cso_2017_ult",
    ProductType.INDEXED_UL: "cso_2017_ult",
    ProductType.VARIABLE_UL: "cso_2017_ult",
}

_MORTALITY_MODE_LABELS: dict[str, str] = {
    "synthetic": "Synthetic (demo, wide age range)",
    "qx_csv": "Static q_x CSV",
    "rp2014_mp2016": "RP-2014 Healthy Male + MP-2016 (xlsx or cached CSV)",
    "us_ssa_2015_period": "US SSA 2015 period life table (sex-specific default for Term)",
    "cso_2017_ult": "2017 CSO Ultimate (sex × smoker, placeholder synthetic CSV)",
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
        selected_info_message="Whole Life (single premium): premium solved as PV of benefits, mortality from CSO 2017 Ultimate placeholder.",
        projection_csv_filename="pricing_projection_whole_life.csv",
        recalc_workbook_filename="whole_life_recalc_model.xlsx",
    ),
    ProductType.VARIABLE_ANNUITY: ProductUIConfig(
        selected_info_message="Variable Annuity (single premium): GMDB = max(AV, premium). Sub-account is deterministic CSV by default; Monte Carlo simulates GBM.",
        projection_csv_filename="pricing_projection_variable_annuity.csv",
        recalc_workbook_filename="variable_annuity_recalc_model.xlsx",
    ),
    ProductType.MYGA: ProductUIConfig(
        selected_info_message="MYGA (multi-year guaranteed annuity): single premium accumulates at the declared rate for the guarantee period.",
        projection_csv_filename="pricing_projection_myga.csv",
        recalc_workbook_filename="myga_recalc_model.xlsx",
    ),
    ProductType.FIA: ProductUIConfig(
        selected_info_message="FIA (fixed indexed annuity): annual point-to-point credit with cap, floor, and participation. Floor 0 by default.",
        projection_csv_filename="pricing_projection_fia.csv",
        recalc_workbook_filename="fia_recalc_model.xlsx",
    ),
    ProductType.UNIVERSAL_LIFE: ProductUIConfig(
        selected_info_message="Universal Life (single premium): monthly cycle of credit -> COI -> expense charge. Type A death benefit.",
        projection_csv_filename="pricing_projection_universal_life.csv",
        recalc_workbook_filename="universal_life_recalc_model.xlsx",
    ),
    ProductType.INDEXED_UL: ProductUIConfig(
        selected_info_message="Indexed UL (IUL): UL mechanics with annual point-to-point crediting on segment anniversaries.",
        projection_csv_filename="pricing_projection_indexed_ul.csv",
        recalc_workbook_filename="indexed_ul_recalc_model.xlsx",
    ),
    ProductType.VARIABLE_UL: ProductUIConfig(
        selected_info_message="Variable UL (VUL): UL mechanics with sub-account return as credit (deterministic CSV or GBM Monte Carlo).",
        projection_csv_filename="pricing_projection_variable_ul.csv",
        recalc_workbook_filename="variable_ul_recalc_model.xlsx",
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
    ProductType.MYGA: _MYGA_ADAPTER,
    ProductType.FIA: _FIA_ADAPTER,
    ProductType.VARIABLE_ANNUITY: _VA_ADAPTER,
    ProductType.WHOLE_LIFE: _WL_ADAPTER,
    ProductType.UNIVERSAL_LIFE: _UL_ADAPTER,
    ProductType.INDEXED_UL: _IUL_ADAPTER,
    ProductType.VARIABLE_UL: _VUL_ADAPTER,
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


def product_adapters_by_type() -> Mapping[ProductType, ProductAdapter]:
    """Read-only legacy adapter registry view.

    New code should prefer ``products.product_adapters_by_type()``, which is
    derived from ``ProductDefinition``. This view remains for compatibility
    checks and migration tests that compare the canonical and legacy wires.
    """
    return MappingProxyType(dict(_PRODUCT_ADAPTERS))


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
        ProductType.MYGA,
        ProductType.FIA,
        ProductType.VARIABLE_ANNUITY,
        ProductType.WHOLE_LIFE,
        ProductType.UNIVERSAL_LIFE,
        ProductType.INDEXED_UL,
        ProductType.VARIABLE_UL,
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


def _accumulation_pricing_metrics(result: Any) -> tuple[PricingMetric, ...]:
    """Shared accumulation-product metrics (MYGA / FIA / VA)."""
    av_end = float(getattr(result, "account_value_end_month", np.array([0.0]))[-1])
    return (
        PricingMetric(
            label="Single premium (input)", value=float(result.single_premium), is_money=True
        ),
        PricingMetric(
            label="PV benefit (death+maturity)", value=float(result.pv_benefit), is_money=True
        ),
        PricingMetric(
            label="PV monthly expenses", value=float(result.pv_monthly_expenses), is_money=True
        ),
        PricingMetric(label="Account value at horizon", value=av_end, is_money=True),
    )


def _life_single_premium_metrics(result: Any) -> tuple[PricingMetric, ...]:
    """Shared life-product metrics (WL / UL / IUL / VUL)."""
    av_end = float(getattr(result, "account_value_end_month", np.array([0.0]))[-1])
    return (
        PricingMetric(label="Single premium", value=float(result.single_premium), is_money=True),
        PricingMetric(
            label="PV claims (face × death-prob)", value=float(result.pv_benefit), is_money=True
        ),
        PricingMetric(
            label="PV monthly expenses", value=float(result.pv_monthly_expenses), is_money=True
        ),
        PricingMetric(label="Account value at horizon", value=av_end, is_money=True),
    )


def _wl_pricing_metrics(result: Any) -> tuple[PricingMetric, ...]:
    """WL has no AV; show face amount instead."""
    face = float(getattr(result, "face_amount", 0.0))
    return (
        PricingMetric(
            label="Single premium (NSP + expenses)",
            value=float(result.single_premium),
            is_money=True,
        ),
        PricingMetric(
            label="PV claims (face × death-prob)", value=float(result.pv_benefit), is_money=True
        ),
        PricingMetric(
            label="PV monthly expenses", value=float(result.pv_monthly_expenses), is_money=True
        ),
        PricingMetric(label="Face amount", value=face, is_money=True),
    )


# Per-product pricing-metric formatters. Adding a new product means one
# additional entry here; the dispatcher stays branch-free.
_PRICING_METRIC_FORMATTERS: dict[ProductType, Callable[[Any], tuple[PricingMetric, ...]]] = {
    ProductType.SPIA: _spia_pricing_metrics,
    ProductType.TERM_LIFE: _term_pricing_metrics,
    ProductType.RILA: _rila_pricing_metrics,
    ProductType.MYGA: _accumulation_pricing_metrics,
    ProductType.FIA: _accumulation_pricing_metrics,
    ProductType.VARIABLE_ANNUITY: _accumulation_pricing_metrics,
    ProductType.WHOLE_LIFE: _wl_pricing_metrics,
    ProductType.UNIVERSAL_LIFE: _life_single_premium_metrics,
    ProductType.INDEXED_UL: _life_single_premium_metrics,
    ProductType.VARIABLE_UL: _life_single_premium_metrics,
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


def pricing_metric_formatters_by_type() -> Mapping[
    ProductType, Callable[[Any], tuple[PricingMetric, ...]]
]:
    """Read-only legacy metric-formatter registry view.

    ProductDefinition-derived views are canonical for migration work; this
    compatibility view lets tests prove the legacy private registry has not
    drifted while existing callers are preserved.
    """
    return MappingProxyType(dict(_PRICING_METRIC_FORMATTERS))


def get_product_ui_config(product_type: ProductType) -> ProductUIConfig:
    return _PRODUCT_UI_CONFIG[product_type]
