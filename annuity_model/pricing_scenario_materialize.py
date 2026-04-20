"""Streamlit-free materialization of :class:`portfolio.RunScenario` from Pricing Run ``run_*`` seeds.

Keeps portfolio / CLI economics aligned with ``pricing_ui`` + ``build_run_form_seed_defaults``.
"""

from __future__ import annotations

import io
from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import Any, Literal

import numpy as np
import pandas as pd

import pricing_projection as sp
from portfolio import PolicyInput, RunScenario
from pricing_run_form_state import RUN_KEY, RUN_STATE_KEY_NAMES, build_run_form_seed_defaults
from product_registry import ProductType
from ssa_2015_period_qx_embedded import SSA_2015_PERIOD_QX_CSV

ANN_MODEL_ROOT = Path(__file__).resolve().parent


def reference_product_type_for_portfolio_scenario(policies: Sequence[PolicyInput]) -> ProductType:
    """Pick ``default_product_type`` for :func:`build_run_form_seed_defaults` / shared mortality.

    RP-2014 annuitant extracts used by SPIA defaults start at age 50. If any policy issues below
    that band (common for Term / UL), use **Term Life** defaults (US SSA 2015 period) so the
    shared :class:`~pricing_projection.MortalityTableQx` covers working ages. Otherwise use the
    first policy's type so single-policy books match Pricing Run row order.
    """
    if not policies:
        return ProductType.SPIA
    min_age = min(int(getattr(p.contract, "issue_age", 999)) for p in policies)
    if min_age < 50:
        return ProductType.TERM_LIFE
    return policies[0].product_type


def resolve_repo_path(path_str: str, *, repo_root: Path | None = None) -> Path:
    path = Path(path_str.strip())
    if path.is_absolute():
        return path
    root = repo_root if repo_root is not None else ANN_MODEL_ROOT
    return (root / path).resolve()


def minimal_mortality() -> sp.MortalityTableQx:
    ages = np.arange(0, 121, dtype=int)
    qx = np.clip(0.005 + ages * 1e-5, 1e-6, 0.4)
    return sp.MortalityTableQx(ages, qx)


def build_yield_curve_from_seeds(
    seeds: Mapping[str, Any],
    *,
    repo_root: Path | None = None,
) -> sp.YieldCurve:
    y_mode = str(seeds.get(RUN_KEY.Y_MODE, "par_bootstrap"))
    flat_rate = float(seeds.get(RUN_KEY.FLAT_RATE, 0.04))
    zero_csv = str(seeds.get(RUN_KEY.ZERO_CSV, sp.DEFAULT_ZERO_CURVE_CSV))
    par_csv = str(seeds.get(RUN_KEY.PAR_CSV, sp.DEFAULT_PAR_CURVE_CSV))
    coupon_freq = int(seeds.get(RUN_KEY.COUPON_FREQ, 2))
    if y_mode == "flat":
        return sp.YieldCurve.from_flat_rate(float(flat_rate))
    if y_mode == "zero_csv":
        return sp.YieldCurve.load_zero_curve_csv(str(resolve_repo_path(zero_csv, repo_root=repo_root)))
    return sp.YieldCurve.load_par_yield_csv_and_bootstrap(
        str(resolve_repo_path(par_csv, repo_root=repo_root)),
        coupon_freq=int(coupon_freq),
    )


def build_mortality_from_seeds(
    seeds: Mapping[str, Any],
    *,
    product_type: ProductType,
    sex: Literal["male", "female"],
    repo_root: Path | None = None,
) -> tuple[sp.MortalityTableQx | sp.MortalityTableRP2014MP2016, bool]:
    mode = str(seeds.get(RUN_KEY.M_MODE, "rp2014_mp2016"))
    qx_csv = str(seeds.get(RUN_KEY.QX_CSV, sp.DEFAULT_MORTALITY_QX_CSV))
    rp_xlsx = str(seeds.get(RUN_KEY.RP_XLSX, sp.DEFAULT_RP2014_XLSX))
    rp_out = str(seeds.get(RUN_KEY.RP_OUT, sp.DEFAULT_RP2014_MALE_HEALTHY_QX_CSV))
    mp_xlsx = str(seeds.get(RUN_KEY.MP_XLSX, sp.DEFAULT_MP2016_XLSX))
    mp_out = str(seeds.get(RUN_KEY.MP_OUT, sp.DEFAULT_MP2016_MALE_IMPROVEMENT_CSV))

    if mode == "us_ssa_2015_period":
        if product_type != ProductType.TERM_LIFE:
            raise ValueError("US SSA 2015 period mortality is currently scoped to Term Life.")
        raw = pd.read_csv(io.StringIO(SSA_2015_PERIOD_QX_CSV))
        qx_col = "male_qx" if sex == "male" else "female_qx"
        return (
            sp.MortalityTableQx(raw["age"].to_numpy(dtype=int), raw[qx_col].to_numpy(dtype=float)),
            False,
        )
    if mode == "synthetic":
        return minimal_mortality(), False
    if mode == "qx_csv":
        return sp.MortalityTableQx.load_qx_csv(str(resolve_repo_path(qx_csv, repo_root=repo_root))), False
    if mode == "cso_2017_ult":
        from mortality_2017_cso import MortalityTable2017CSO

        cso = MortalityTable2017CSO.load(sex=sex, smoker_class="nonsmoker")
        return cso.table, False
    base_qx = sp.ensure_rp2014_male_healthy_annuitant_qx_csv(
        rp2014_xlsx_path=str(resolve_repo_path(rp_xlsx, repo_root=repo_root)),
        out_csv_path=str(resolve_repo_path(rp_out, repo_root=repo_root)),
    )
    mp_ages, mp_years, mp_i = sp.ensure_mp2016_male_improvement_csv(
        mp2016_xlsx_path=str(resolve_repo_path(mp_xlsx, repo_root=repo_root)),
        out_csv_path=str(resolve_repo_path(mp_out, repo_root=repo_root)),
    )
    mortality = sp.MortalityTableRP2014MP2016(
        base_qx_2014=base_qx,
        mp2016_ages=mp_ages,
        mp2016_years=mp_years,
        mp2016_i_matrix=mp_i,
        base_year=2014,
    )
    return mortality, True


def merged_run_form_seeds(
    session: Mapping[str, Any],
    *,
    default_product_type: ProductType,
    saved_inputs: Mapping[str, Any] | None = None,
    meta: Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    """Merge ``build_run_form_seed_defaults`` with any ``run_*`` keys present in *session*."""
    saved = dict(saved_inputs or session.get("pricing_run_inputs") or {})
    m = dict(meta or session.get("pricing_meta") or {})
    base = build_run_form_seed_defaults(
        product_default=default_product_type.value,
        saved_inputs=saved,
        meta=m,
        default_product_type=default_product_type,
    )
    for k in RUN_STATE_KEY_NAMES:
        if k in session:
            base[k] = session[k]
    return base


def run_scenario_from_pricing_seeds(
    seeds: Mapping[str, Any],
    *,
    default_product_type: ProductType,
    sex: Literal["male", "female"],
    repo_root: Path | None = None,
) -> RunScenario:
    """Build :class:`RunScenario` exactly like Pricing Run ``adapter.price`` economics (deterministic)."""
    yc = build_yield_curve_from_seeds(seeds, repo_root=repo_root)
    mort, needs_vy = build_mortality_from_seeds(
        seeds, product_type=default_product_type, sex=sex, repo_root=repo_root
    )
    vy = int(seeds[RUN_KEY.VALUATION_YEAR]) if needs_vy else None
    horizon_age = int(seeds[RUN_KEY.HORIZON_AGE])
    spread = float(seeds[RUN_KEY.SPREAD])
    expense_mode = str(seeds.get(RUN_KEY.EXPENSE_MODE, "csv"))
    expenses_csv = str(seeds.get(RUN_KEY.EXPENSES_CSV, sp.DEFAULT_EXPENSES_CSV))
    expenses_arg: sp.ExpenseAssumptions | None = None
    if expense_mode == "manual":
        expenses_arg = sp.ExpenseAssumptions(
            policy_expense_dollars=float(seeds.get(RUN_KEY.POLICY_EXPENSE, 0.0)),
            premium_expense_rate=float(seeds.get(RUN_KEY.PREMIUM_EXPENSE_PCT, 0.0)) / 100.0,
            monthly_expense_dollars=float(seeds.get(RUN_KEY.MONTHLY_EXPENSE, 0.0)),
        )
    use_index = bool(seeds.get(RUN_KEY.USE_INDEX, True))
    index_csv = str(seeds.get(RUN_KEY.INDEX_CSV, sp.DEFAULT_SP500_SCENARIO_CSV))
    idx_path = str(resolve_repo_path(index_csv, repo_root=repo_root)) if use_index else None
    expense_annual_inflation = float(seeds.get(RUN_KEY.EXPENSE_INFLATION_PCT, 2.5)) / 100.0
    return RunScenario(
        yield_curve=yc,
        mortality=mort,
        horizon_age=horizon_age,
        spread=spread,
        valuation_year=vy,
        expenses=expenses_arg,
        expenses_csv_path=str(resolve_repo_path(expenses_csv, repo_root=repo_root)),
        index_scenario_csv_path=idx_path,
        expense_annual_inflation=expense_annual_inflation,
    )


def cli_default_run_scenario(
    *,
    default_product_type: ProductType = ProductType.SPIA,
    sex: Literal["male", "female"] = "male",
    repo_root: Path | None = None,
) -> RunScenario:
    """CLI / smoke: first-paint Pricing Run economics without Streamlit session."""
    seeds = build_run_form_seed_defaults(
        product_default=default_product_type.value,
        saved_inputs={},
        meta={},
        default_product_type=default_product_type,
    )
    return run_scenario_from_pricing_seeds(
        seeds, default_product_type=default_product_type, sex=sex, repo_root=repo_root
    )


def run_scenario_for_portfolio_policies(
    session: Mapping[str, Any],
    policies: Sequence[PolicyInput],
    *,
    sex: Literal["male", "female"],
    repo_root: Path | None = None,
) -> RunScenario:
    """Merge Pricing Run session economics with mortality keys pinned to the portfolio reference product."""
    ref = reference_product_type_for_portfolio_scenario(policies)
    ref_base = build_run_form_seed_defaults(
        product_default=ref.value,
        saved_inputs={},
        meta={},
        default_product_type=ref,
    )
    seeds = merged_run_form_seeds(session, default_product_type=ref)
    for mk in (
        RUN_KEY.M_MODE,
        RUN_KEY.QX_CSV,
        RUN_KEY.RP_XLSX,
        RUN_KEY.RP_OUT,
        RUN_KEY.MP_XLSX,
        RUN_KEY.MP_OUT,
    ):
        seeds[mk] = ref_base[mk]
    return run_scenario_from_pricing_seeds(seeds, default_product_type=ref, sex=sex, repo_root=repo_root)


__all__ = [
    "ANN_MODEL_ROOT",
    "build_mortality_from_seeds",
    "build_yield_curve_from_seeds",
    "cli_default_run_scenario",
    "merged_run_form_seeds",
    "minimal_mortality",
    "reference_product_type_for_portfolio_scenario",
    "resolve_repo_path",
    "run_scenario_for_portfolio_policies",
    "run_scenario_from_pricing_seeds",
]
