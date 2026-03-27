from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from typing import Any, Protocol

import numpy as np

import pricing_projection as sp
import term_projection as tp
from build_pricing_excel_workbook import ExcelBuildSpec, excel_spec_from_launcher
from build_term_excel_workbook import TermExcelBuildSpec, term_excel_spec_from_launcher


class ProductType(str, Enum):
    SPIA = "spia"
    TERM_LIFE = "term_life"
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
    @property
    def product_type(self) -> ProductType: ...

    @property
    def display_name(self) -> str: ...

    def is_available(self) -> bool: ...

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
    ) -> object: ...

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
    ) -> object: ...

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
    ) -> ExcelBuildSpec | TermExcelBuildSpec: ...


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

_PRODUCT_DISPLAY_NAME: dict[ProductType, str] = {
    ProductType.SPIA: "SPIA",
    ProductType.TERM_LIFE: "Term Life (20Y)",
    ProductType.WHOLE_LIFE: "Whole Life (coming soon)",
    ProductType.VARIABLE_ANNUITY: "Variable Annuity (coming soon)",
}

_PRODUCT_CAPABILITIES: dict[ProductType, ProductCapabilities] = {
    ProductType.SPIA: ProductCapabilities(supports_economic_scenario=True, supports_monte_carlo=True),
    ProductType.TERM_LIFE: ProductCapabilities(supports_economic_scenario=False, supports_monte_carlo=False),
    ProductType.WHOLE_LIFE: ProductCapabilities(supports_economic_scenario=False, supports_monte_carlo=False),
    ProductType.VARIABLE_ANNUITY: ProductCapabilities(supports_economic_scenario=True, supports_monte_carlo=True),
}

_PRODUCT_MORTALITY_MODE_OPTIONS: dict[ProductType, tuple[str, ...]] = {
    ProductType.SPIA: ("synthetic", "qx_csv", "rp2014_mp2016"),
    ProductType.TERM_LIFE: ("us_ssa_2015_period", "qx_csv", "synthetic"),
    ProductType.WHOLE_LIFE: ("synthetic", "qx_csv"),
    ProductType.VARIABLE_ANNUITY: ("synthetic", "qx_csv", "rp2014_mp2016"),
}

_PRODUCT_DEFAULT_MORTALITY_MODE: dict[ProductType, str] = {
    ProductType.SPIA: "rp2014_mp2016",
    ProductType.TERM_LIFE: "us_ssa_2015_period",
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
    term_length_options=("20 years",),
    premium_mode_options=("Level monthly",),
    benefit_timing_options=("EOY death benefit",),
    default_monthly_premium=250.0,
)

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


def get_product_adapter(product_type: ProductType) -> ProductAdapter:
    if product_type == ProductType.SPIA:
        return _SPIA_ADAPTER
    if product_type == ProductType.TERM_LIFE:
        return _TERM_ADAPTER
    raise NotImplementedError(f"{_PRODUCT_DISPLAY_NAME[product_type]} is not implemented yet.")


def product_options_for_ui() -> list[ProductType]:
    return [ProductType.SPIA, ProductType.TERM_LIFE, ProductType.WHOLE_LIFE, ProductType.VARIABLE_ANNUITY]


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


def get_pricing_metrics(product_type: ProductType, result: Any) -> tuple[PricingMetric, ...]:
    if product_type == ProductType.TERM_LIFE:
        pv_claims = float(getattr(result, "pv_benefit"))
        pv_premiums = float(-float(getattr(result, "pv_monthly_expenses")))
        net_pv = float(getattr(result, "single_premium"))
        economic_reserve = np.asarray(getattr(result, "economic_reserve", np.asarray([], dtype=float)))
        issue_reserve = float(economic_reserve[0]) if economic_reserve.size else float("nan")
        return (
            PricingMetric(label="PV claims", value=pv_claims, is_money=True),
            PricingMetric(label="PV premiums", value=pv_premiums, is_money=True),
            PricingMetric(label="Net PV (claims - premiums)", value=net_pv, is_money=True),
            PricingMetric(label="Issue reserve", value=issue_reserve, is_money=True),
        )
    return (
        PricingMetric(label="Single premium", value=float(getattr(result, "single_premium")), is_money=True),
        PricingMetric(label="PV benefit", value=float(getattr(result, "pv_benefit")), is_money=True),
        PricingMetric(label="PV monthly expenses", value=float(getattr(result, "pv_monthly_expenses")), is_money=True),
        PricingMetric(label="Annuity factor", value=float(getattr(result, "annuity_factor")), is_money=False),
    )


def get_product_ui_config(product_type: ProductType) -> ProductUIConfig:
    return _PRODUCT_UI_CONFIG[product_type]
