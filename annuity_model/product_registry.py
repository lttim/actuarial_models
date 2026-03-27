from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from typing import Protocol

import numpy as np

import spia_projection as sp
from build_spia_excel_workbook import ExcelBuildSpec, excel_spec_from_launcher


class ProductType(str, Enum):
    SPIA = "spia"
    WHOLE_LIFE = "whole_life"
    VARIABLE_ANNUITY = "variable_annuity"


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
    ) -> ExcelBuildSpec: ...


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

_PRODUCT_DISPLAY_NAME: dict[ProductType, str] = {
    ProductType.SPIA: "SPIA",
    ProductType.WHOLE_LIFE: "Whole Life (coming soon)",
    ProductType.VARIABLE_ANNUITY: "Variable Annuity (coming soon)",
}


def get_product_adapter(product_type: ProductType) -> ProductAdapter:
    if product_type == ProductType.SPIA:
        return _SPIA_ADAPTER
    raise NotImplementedError(f"{_PRODUCT_DISPLAY_NAME[product_type]} is not implemented yet.")


def product_options_for_ui() -> list[ProductType]:
    return [ProductType.SPIA, ProductType.WHOLE_LIFE, ProductType.VARIABLE_ANNUITY]


def product_label(product_type: ProductType) -> str:
    return _PRODUCT_DISPLAY_NAME[product_type]
