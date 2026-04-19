"""Row-level inforce parsers (CSV columns) -> :data:`product_registry.ProductContract`."""

from __future__ import annotations

from typing import Any, Literal, Mapping, cast

import pandas as pd

import fia_projection as fp
import iul_projection as iulp
import myga_projection as my
import pricing_projection as sp
import rila_projection as rp
import term_projection as tp
import ul_projection as ul
import va_projection as va
import vul_projection as vul
import wl_projection as wl
from product_registry import ProductContract, ProductType


def _sex(val: Any) -> Literal["male", "female"]:
    s = str(val).strip().lower()
    if s not in ("male", "female"):
        raise ValueError(f"sex must be 'male' or 'female', got {val!r}")
    return s  # type: ignore[return-value]


def _req(row: Mapping[str, Any], key: str) -> Any:
    if key not in row:
        raise ValueError(f"missing required column {key!r}")
    v = row[key]
    if v is None or pd.isna(v):
        raise ValueError(f"missing required column {key!r}")
    return v


def _opt_float(row: Mapping[str, Any], key: str, default: float) -> float:
    v = row.get(key)
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return default
    return float(v)


def spia_row_to_contract(row: Mapping[str, Any]) -> ProductContract:
    return sp.SPIAContract(
        issue_age=int(_req(row, "issue_age")),
        sex=_sex(_req(row, "sex")),
        benefit_annual=float(_req(row, "benefit_annual")),
    )


def term_row_to_contract(row: Mapping[str, Any]) -> ProductContract:
    return tp.TermLifeContract(
        issue_age=int(_req(row, "issue_age")),
        sex=_sex(_req(row, "sex")),
        death_benefit=float(_req(row, "death_benefit")),
        monthly_premium=_opt_float(row, "monthly_premium", 250.0),
        term_years=int(_opt_float(row, "term_years", 20.0)),
    )


def rila_row_to_contract(row: Mapping[str, Any]) -> ProductContract:
    return rp.RILAContract(
        issue_age=int(_req(row, "issue_age")),
        sex=_sex(_req(row, "sex")),
        participation=float(_req(row, "participation")),
        cap=float(_req(row, "cap")),
        floor=float(_req(row, "floor")),
        rider_fee_annual=float(_req(row, "rider_fee_annual")),
        segment_months=int(_opt_float(row, "segment_months", 12.0)),
    )


def myga_row_to_contract(row: Mapping[str, Any]) -> ProductContract:
    return my.MYGAContract(
        issue_age=int(_req(row, "issue_age")),
        sex=_sex(_req(row, "sex")),
        single_premium=float(_req(row, "single_premium")),
        declared_rate_annual=float(_req(row, "declared_rate_annual")),
        guarantee_years=int(_opt_float(row, "guarantee_years", 5.0)),
    )


def fia_row_to_contract(row: Mapping[str, Any]) -> ProductContract:
    return fp.FIAContract(
        issue_age=int(_req(row, "issue_age")),
        sex=_sex(_req(row, "sex")),
        single_premium=float(_req(row, "single_premium")),
        participation=float(_req(row, "participation")),
        cap=float(_req(row, "cap")),
        floor=float(_req(row, "floor")),
        rider_fee_annual=_opt_float(row, "rider_fee_annual", 0.0),
        horizon_years=int(_opt_float(row, "horizon_years", 10.0)),
    )


def va_row_to_contract(row: Mapping[str, Any]) -> ProductContract:
    basis_raw = row.get("gmdb_basis", "return_of_premium")
    basis = str(basis_raw).strip() if basis_raw is not None and not pd.isna(basis_raw) else "return_of_premium"
    if basis not in ("return_of_premium", "max_anniversary"):
        basis = "return_of_premium"
    return va.VAContract(
        issue_age=int(_req(row, "issue_age")),
        sex=_sex(_req(row, "sex")),
        single_premium=float(_req(row, "single_premium")),
        me_charge_annual=_opt_float(row, "me_charge_annual", 0.014),
        gmdb_basis=cast(Literal["return_of_premium", "max_anniversary"], basis),
        horizon_years=int(_opt_float(row, "horizon_years", 20.0)),
    )


def wl_row_to_contract(row: Mapping[str, Any]) -> ProductContract:
    return wl.WLContract(
        issue_age=int(_req(row, "issue_age")),
        sex=_sex(_req(row, "sex")),
        face_amount=float(_req(row, "face_amount")),
    )


def ul_row_to_contract(row: Mapping[str, Any]) -> ProductContract:
    return ul.ULContract(
        issue_age=int(_req(row, "issue_age")),
        sex=_sex(_req(row, "sex")),
        face_amount=float(_req(row, "face_amount")),
        single_premium=float(_req(row, "single_premium")),
    )


def iul_row_to_contract(row: Mapping[str, Any]) -> ProductContract:
    return iulp.IULContract(
        issue_age=int(_req(row, "issue_age")),
        sex=_sex(_req(row, "sex")),
        face_amount=float(_req(row, "face_amount")),
        single_premium=float(_req(row, "single_premium")),
        participation=float(_req(row, "participation")),
        cap=float(_req(row, "cap")),
        floor=float(_req(row, "floor")),
    )


def vul_row_to_contract(row: Mapping[str, Any]) -> ProductContract:
    return vul.VULContract(
        issue_age=int(_req(row, "issue_age")),
        sex=_sex(_req(row, "sex")),
        face_amount=float(_req(row, "face_amount")),
        single_premium=float(_req(row, "single_premium")),
        subaccount_drift_annual=_opt_float(row, "subaccount_drift_annual", 0.06),
        subaccount_vol_annual=_opt_float(row, "subaccount_vol_annual", 0.15),
    )


ROW_TO_CONTRACT: dict[ProductType, Any] = {
    ProductType.SPIA: spia_row_to_contract,
    ProductType.TERM_LIFE: term_row_to_contract,
    ProductType.RILA: rila_row_to_contract,
    ProductType.MYGA: myga_row_to_contract,
    ProductType.FIA: fia_row_to_contract,
    ProductType.VARIABLE_ANNUITY: va_row_to_contract,
    ProductType.WHOLE_LIFE: wl_row_to_contract,
    ProductType.UNIVERSAL_LIFE: ul_row_to_contract,
    ProductType.INDEXED_UL: iul_row_to_contract,
    ProductType.VARIABLE_UL: vul_row_to_contract,
}


def contract_from_inforce_row(row: Mapping[str, Any]) -> ProductContract:
    raw = _req(row, "product_type")
    pt = ProductType(str(raw).strip())
    fn = ROW_TO_CONTRACT.get(pt)
    if fn is None:
        raise ValueError(f"unsupported product_type for inforce row: {raw!r}")
    return fn(row)


__all__ = [
    "ROW_TO_CONTRACT",
    "contract_from_inforce_row",
    "fia_row_to_contract",
    "iul_row_to_contract",
    "myga_row_to_contract",
    "rila_row_to_contract",
    "spia_row_to_contract",
    "term_row_to_contract",
    "ul_row_to_contract",
    "va_row_to_contract",
    "vul_row_to_contract",
    "wl_row_to_contract",
]
