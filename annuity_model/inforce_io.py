"""Load :class:`portfolio.PolicyInput` rows from CSV or Excel inforce files."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from inforce_parsers import contract_from_inforce_row
from portfolio import PolicyInput
from product_registry import ProductType


def load_policy_inputs_from_csv(path: str | Path) -> tuple[PolicyInput, ...]:
    """Parse a CSV with at least ``product_type`` and product-specific columns."""
    df = pd.read_csv(Path(path))
    return load_policy_inputs_from_csv_from_dataframe(df)


def load_policy_inputs_from_excel(path: str | Path, *, sheet_name: str | int = 0) -> tuple[PolicyInput, ...]:
    """Parse the first sheet (or *sheet_name*) of an Excel inforce workbook."""
    df = pd.read_excel(Path(path), sheet_name=sheet_name)
    return load_policy_inputs_from_csv_from_dataframe(df)


def load_policy_inputs_from_csv_from_dataframe(df: pd.DataFrame) -> tuple[PolicyInput, ...]:
    """Shared parser body for CSV / Excel paths."""
    if "product_type" not in df.columns:
        raise ValueError("inforce table must include a 'product_type' column.")
    out: list[PolicyInput] = []
    for _, row in df.iterrows():
        row_d = {str(k): row[k] for k in df.columns}
        pid = row_d.get("policy_id")
        if pid is None or (isinstance(pid, float) and pd.isna(pid)):
            policy_id = None
        else:
            policy_id = str(pid)
        contract = contract_from_inforce_row(row_d)
        pt = ProductType(str(row_d["product_type"]).strip())
        out.append(PolicyInput(product_type=pt, contract=contract, policy_id=policy_id))
    return tuple(out)


__all__ = [
    "load_policy_inputs_from_csv",
    "load_policy_inputs_from_csv_from_dataframe",
    "load_policy_inputs_from_excel",
]
