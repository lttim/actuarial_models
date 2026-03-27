"""
Unified Streamlit workspace for the SPIA model: overview, configurable pricing run,
interactive charts, and embedded unit-test dashboard.

Run from the annuity_model folder:
    streamlit run pricing_ui.py
Or: run_pricing_ui.bat (Windows).
"""

from __future__ import annotations

import io
import os
import sys
from pathlib import Path
from typing import Any, Literal, MutableMapping

import dataclasses
import datetime as _dt
import json
import time

os.environ.setdefault("MPLBACKEND", "Agg")

import altair as alt
import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import pricing_projection as sp
import term_projection as tp
from alm_excel_ladder import ALM_ENGINE_SHEET

from build_pricing_excel_workbook import (
    ALM_ENGINE_FIELD_GUIDE_SHEET,
    ALM_ENGINE_STEP_MONTHS,
    ALM_EXCEL_PATH_MONTH_CAP,
    ALMExcelSnapshot,
    ALM_PROJECTION_FIRST_DATA_ROW,
    ALM_SHEET_NAME,
    ExcelPythonSnapshot,
    LIABILITY_SHEET_NAME,
    MCExcelSnapshot,
    alm_excel_downsample_snapshot,
    alm_excel_snapshot_from_result,
    alm_excel_truncate_snapshot,
    mc_excel_snapshot_from_result,
)
from product_excel import build_product_workbook
from product_registry import (
    ProductType,
    get_product_adapter,
    get_product_capabilities,
    get_product_default_mortality_mode,
    get_mortality_mode_label,
    get_product_ui_config,
    get_pricing_metrics,
    get_product_mortality_mode_options,
    get_term_contract_ui_config,
    product_label,
    product_options_for_ui,
)
from test_dashboard import render_unit_tests_page


def _maybe_alm_excel_snapshot_for_workbook() -> ALMExcelSnapshot | None:
    alm = st.session_state.get("alm_last")
    asm = st.session_state.get("alm_last_assumptions")
    if not isinstance(alm, sp.ALMResult) or not isinstance(asm, sp.ALMAssumptions):
        return None
    if st.session_state.get("alm_last_pricing_run_id") != st.session_state.get("pricing_run_id"):
        return None
    aum_tag = st.session_state.get("alm_last_initial_asset_market_value")
    return alm_excel_snapshot_from_result(
        alm,
        asm,
        initial_asset_market_value=float(aum_tag) if aum_tag is not None else None,
    )


def _refresh_pricing_excel_workbook_in_session() -> None:
    """Rebuild `pricing_xlsx_bytes` from the current pricing result and optional MC/ALM session state."""
    res = st.session_state.get("pricing_res")
    contract = st.session_state.get("pricing_contract")
    ctx = st.session_state.get("pricing_excel_context") or {}
    if res is None or contract is None:
        return
    yc = ctx.get("yield_curve")
    mort = ctx.get("mortality")
    if not isinstance(yc, sp.YieldCurve) or not isinstance(
        mort, (sp.MortalityTableQx, sp.MortalityTableRP2014MP2016)
    ):
        return
    expenses = ctx.get("expenses")
    if not isinstance(expenses, sp.ExpenseAssumptions):
        return
    meta = st.session_state.get("pricing_meta") or {}
    product_raw = st.session_state.get("pricing_product_type", ProductType.SPIA.value)
    try:
        product_type = ProductType(str(product_raw))
    except ValueError:
        product_type = ProductType.SPIA
    adapter = get_product_adapter(product_type)
    vy_raw = ctx.get("valuation_year")
    vy = int(vy_raw) if vy_raw is not None else 2025
    mc_snap: MCExcelSnapshot | None = None
    mc = st.session_state.get("pricing_mc")
    mc_params = st.session_state.get("pricing_mc_params") or {}
    if mc is not None:
        mc_snap = mc_excel_snapshot_from_result(
            mc,
            annual_drift=float(mc_params.get("annual_drift", 0.06)),
            annual_vol=float(mc_params.get("annual_vol", 0.15)),
            s0=float(mc_params.get("s0", 100.0)),
        )
    alm_snap = _maybe_alm_excel_snapshot_for_workbook()
    alm_asm = st.session_state.get("alm_last_assumptions")
    try:
        spec = adapter.excel_spec_from_run(
            contract=contract,
            yield_curve=yc,
            mortality=mort,
            horizon_age=int(ctx.get("horizon_age", 110)),
            spread=float(ctx.get("spread", 0.0)),
            valuation_year=vy,
            expenses=expenses,
            yield_mode_label=str(meta.get("yield_mode", "")),
            mortality_mode_label=str(meta.get("mortality_mode", "")),
            expense_mode_label=str(meta.get("expense_mode", "")),
            index_s0=float(res.index_s0),
            index_levels_at_payment=res.index_level_at_payment,
            expense_annual_inflation=float(res.expense_annual_inflation),
        )
        st.session_state["pricing_xlsx_bytes"] = build_product_workbook(
            product_type=product_type,
            spec=spec,
            out_path=None,
            python_snapshot=ExcelPythonSnapshot(
                pv_benefit=float(res.pv_benefit),
                pv_monthly_expenses=float(res.pv_monthly_expenses),
                pv_monthly_total=float(res.pv_benefit + res.pv_monthly_expenses),
                single_premium=float(res.single_premium),
                annuity_factor=float(res.annuity_factor),
            ),
            mc_snapshot=mc_snap,
            alm_snapshot=alm_snap,
            alm_assumptions=alm_asm if isinstance(alm_asm, sp.ALMAssumptions) else None,
        )
        st.session_state["pricing_xlsx_has_mc"] = mc_snap is not None
        st.session_state["pricing_xlsx_has_alm"] = alm_snap is not None
        st.session_state.pop("pricing_xlsx_built_error", None)
    except Exception as ex:
        st.session_state["pricing_xlsx_bytes"] = None
        st.session_state.pop("pricing_xlsx_has_mc", None)
        st.session_state.pop("pricing_xlsx_has_alm", None)
        st.session_state["pricing_xlsx_built_error"] = repr(ex)


def _ensure_excel_workbook_includes_current_alm() -> None:
    """If ALM completed after the last Excel build, regenerate the workbook so download includes ALM_Projection."""
    if not isinstance(st.session_state.get("pricing_xlsx_bytes"), bytes):
        return
    want_alm = _maybe_alm_excel_snapshot_for_workbook() is not None
    has_alm = bool(st.session_state.get("pricing_xlsx_has_alm", False))
    if want_alm != has_alm:
        _refresh_pricing_excel_workbook_in_session()


def _resolve_path(p: str) -> Path:
    path = Path(p.strip())
    if path.is_absolute():
        return path
    return (ROOT / path).resolve()


def _minimal_mortality() -> sp.MortalityTableQx:
    ages = np.arange(0, 121, dtype=int)
    qx = np.clip(0.005 + ages * 1e-5, 1e-6, 0.4)
    return sp.MortalityTableQx(ages, qx)


def _round_for_visuals(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    numeric_cols = out.select_dtypes(include=[np.number]).columns
    out.loc[:, numeric_cols] = out.loc[:, numeric_cols].round(0)
    return out


def _alm_surplus_chart(ages: np.ndarray | pd.Series, surplus: np.ndarray | pd.Series) -> None:
    """Surplus vs attained age with a y = 0 reference line (above / below zero)."""
    df = pd.DataFrame(
        {
            "Attained age": np.asarray(ages, dtype=float),
            "Surplus": np.asarray(surplus, dtype=float),
        }
    )
    line = (
        alt.Chart(df)
        .mark_line()
        .encode(
            x=alt.X("Attained age:Q", title="Attained age"),
            y=alt.Y("Surplus:Q", title="Surplus ($)"),
        )
    )
    rule = (
        alt.Chart(pd.DataFrame({"y": [0.0]}))
        .mark_rule(color="#888", strokeDash=[4, 4])
        .encode(y="y:Q")
    )
    layered = (
        (line + rule)
        .properties(
            title="Surplus (asset market value minus liability PV)",
            height=320,
        )
        .resolve_scale(y="shared")
    )
    st.altair_chart(layered.interactive(), use_container_width=True)


def _number_cols_no_decimals(df: pd.DataFrame) -> dict[str, st.column_config.NumberColumn]:
    numeric_cols = list(df.select_dtypes(include=[np.number]).columns)
    return {c: st.column_config.NumberColumn(format="%,.0f") for c in numeric_cols}


def _pop_session_keys(keys: list[str]) -> None:
    for k in keys:
        st.session_state.pop(k, None)


def _invalidate_diagnostics_export() -> None:
    _pop_session_keys(["diagnostics_json_bytes", "diagnostics_json_filename"])


def _clear_dependent_state_on_pricing_change() -> None:
    # ALM and What-if artifacts are pricing-baseline dependent; clear on pricing changes.
    _pop_session_keys(
        [
            "alm_last",
            "alm_last_assumptions",
            "alm_last_initial_asset_market_value",
            "alm_current_assumptions",
            "alm_current_initial_asset_market_value",
            "whatif_last_params",
            "whatif_last_base_res",
            "whatif_last_shocked_res",
            "whatif_last_baseline_mc",
            "whatif_last_shocked_mc",
            "whatif_last_shocked_curve",
            "whatif_last_shocked_mortality",
            "whatif_last_alm_base",
            "whatif_last_alm_after",
            "whatif_last_alm_assumptions",
            "whatif_last_pricing_run_id",
            "alm_last_pricing_run_id",
        ]
    )
    st.session_state.pop("what_if_mc_cache", None)
    _invalidate_diagnostics_export()


def _serialize_array(arr: Any, *, include_full: bool, max_points: int = 250) -> Any:
    """Serialize numpy-like arrays for JSON with optional truncation."""
    if arr is None:
        return None
    a = np.asarray(arr)
    if include_full or a.size <= max_points:
        return a.tolist()
    # Keep file sizes manageable while still giving a shape + endpoints.
    head_n = min(10, a.size)
    tail_n = min(10, a.size)
    return {
        "truncated": True,
        "len": int(a.size),
        "shape": list(a.shape),
        "head": a[:head_n].tolist(),
        "tail": a[-tail_n:].tolist(),
    }


def _contract_to_dict(contract: sp.SPIAContract) -> dict[str, Any]:
    return {
        "issue_age": int(contract.issue_age),
        "sex": str(contract.sex),
        "benefit_annual": float(contract.benefit_annual),
        "benefit_timing": str(getattr(contract, "benefit_timing", "")),
        "payment_freq_per_year": int(getattr(contract, "payment_freq_per_year", 1)),
        "payment_cessation": str(getattr(contract, "payment_cessation", "")),
    }


def _yield_curve_to_dict(yc: sp.YieldCurve) -> dict[str, Any]:
    return {
        "maturities_years": _serialize_array(yc.maturities_years, include_full=True),
        "zero_rates": _serialize_array(yc.zero_rates, include_full=True),
    }


def _mortality_to_dict(mort: Any) -> dict[str, Any]:
    out: dict[str, Any] = {}
    if hasattr(mort, "ages"):
        out["ages"] = _serialize_array(getattr(mort, "ages"), include_full=True)
    if hasattr(mort, "qx"):
        out["qx"] = _serialize_array(getattr(mort, "qx"), include_full=True)
    if hasattr(mort, "qx_at_int_age"):
        out["qx_at_int_age"] = _serialize_array(getattr(mort, "qx_at_int_age"), include_full=True)
    out["type"] = type(mort).__name__
    return out


def _pricing_result_to_dict(
    res: sp.SPIAProjectionResult,
    contract_state: sp.SPIAContract,
    *,
    include_full: bool,
) -> dict[str, Any]:
    return {
        "contract": _contract_to_dict(contract_state),
        "single_premium": float(res.single_premium),
        "pv_benefit": float(res.pv_benefit),
        "pv_monthly_expenses": float(res.pv_monthly_expenses),
        "annuity_factor": float(res.annuity_factor),
        "times_years": _serialize_array(res.times_years, include_full=include_full),
        "months": _serialize_array(res.months, include_full=include_full),
        "expected_total_cashflows": _serialize_array(res.expected_total_cashflows, include_full=include_full),
        "economic_reserve": _serialize_array(res.economic_reserve, include_full=include_full),
        "survival_to_payment": _serialize_array(res.survival_to_payment, include_full=include_full),
        # Index / inflation scaffolding (needed for full Before/After diagnostics).
        "index_s0": float(res.index_s0),
        "index_level_at_payment": _serialize_array(res.index_level_at_payment, include_full=include_full),
        "index_simple_return": _serialize_array(res.index_simple_return, include_full=include_full),
        "index_log_return": _serialize_array(res.index_log_return, include_full=include_full),
        "index_cumulative_return": _serialize_array(res.index_cumulative_return, include_full=include_full),
    }


def _alm_result_to_dict(alm: sp.ALMResult, asm: sp.ALMAssumptions | None, *, include_buckets: bool, include_full: bool) -> dict[str, Any]:
    out: dict[str, Any] = {
        "assumptions": None,
        "month_index": _serialize_array(alm.month_index, include_full=True),
        "times_years": _serialize_array(alm.times_years, include_full=include_full),
        "asset_market_value": _serialize_array(alm.asset_market_value, include_full=include_full),
        "liability_pv": _serialize_array(alm.liability_pv, include_full=include_full),
        "surplus": _serialize_array(alm.surplus, include_full=include_full),
        "funding_ratio": _serialize_array(alm.funding_ratio, include_full=include_full),
        "liquidity_buffer_months": _serialize_array(alm.liquidity_buffer_months, include_full=include_full),
        "borrowing_balance": _serialize_array(alm.borrowing_balance, include_full=include_full),
        "pv01_assets": float(alm.pv01_assets),
        "pv01_liabilities": float(alm.pv01_liabilities),
        "pv01_net": float(alm.pv01_net),
        "duration_assets_mac": float(alm.duration_assets_mac),
        "duration_liabilities_mac": float(alm.duration_liabilities_mac),
        "duration_gap": float(alm.duration_gap),
    }
    if asm is not None:
        out["assumptions"] = _alm_assumptions_to_dict(asm)
    if include_buckets:
        out["bucket_asset_mv"] = _serialize_array(alm.bucket_asset_mv, include_full=True)
    else:
        out["bucket_asset_mv"] = {
            "shape": list(alm.bucket_asset_mv.shape),
        }
    return out


def _alm_assumptions_to_dict(asm: sp.ALMAssumptions) -> dict[str, Any]:
    return {
        "rebalance_band": float(asm.rebalance_band),
        "rebalance_frequency_months": int(asm.rebalance_frequency_months),
        "reinvest_rule": str(asm.reinvest_rule),
        "disinvest_rule": str(asm.disinvest_rule),
        "rebalance_policy": str(asm.rebalance_policy),
        "borrowing_policy": str(asm.borrowing_policy),
        "borrowing_rate_mode": str(asm.borrowing_rate_mode),
        "borrowing_rate_tenor_years": float(asm.borrowing_rate_tenor_years),
        "borrowing_spread_annual": float(asm.borrowing_spread_annual),
        "borrowing_rate_annual": float(asm.borrowing_rate_annual),
        "liquidity_near_liquid_years": float(asm.liquidity_near_liquid_years),
        "allocation": {
            "buckets": [{"name": b.name, "tenor_years": float(b.tenor_years)} for b in asm.allocation.buckets],
            "weights": _serialize_array(asm.allocation.weights, include_full=True),
        },
    }


def _whatif_result_to_dict(
    *,
    base_res: sp.SPIAProjectionResult,
    shocked_res: sp.SPIAProjectionResult,
    baseline_mc: Any,
    shocked_mc: Any,
    whatif_params: dict[str, Any],
    alm_base: sp.ALMResult | None,
    alm_after: sp.ALMResult | None,
    asm: sp.ALMAssumptions | None,
    include_full: bool,
) -> dict[str, Any]:
    out: dict[str, Any] = {
        "whatif_params": whatif_params,
        "base": {
            "single_premium": float(base_res.single_premium),
            "pv_benefit": float(base_res.pv_benefit),
            "pv_monthly_expenses": float(base_res.pv_monthly_expenses),
            "economic_reserve_issue": float(base_res.economic_reserve[0]) if base_res.economic_reserve.size else None,
            "times_years": _serialize_array(base_res.times_years, include_full=include_full),
            "economic_reserve": _serialize_array(base_res.economic_reserve, include_full=include_full),
            "index_s0": float(base_res.index_s0),
            "index_level_at_payment": _serialize_array(base_res.index_level_at_payment, include_full=include_full),
            "index_simple_return": _serialize_array(base_res.index_simple_return, include_full=include_full),
            "index_log_return": _serialize_array(base_res.index_log_return, include_full=include_full),
            "index_cumulative_return": _serialize_array(base_res.index_cumulative_return, include_full=include_full),
        },
        "after": {
            "single_premium": float(shocked_res.single_premium),
            "pv_benefit": float(shocked_res.pv_benefit),
            "pv_monthly_expenses": float(shocked_res.pv_monthly_expenses),
            "economic_reserve_issue": float(shocked_res.economic_reserve[0]) if shocked_res.economic_reserve.size else None,
            "times_years": _serialize_array(shocked_res.times_years, include_full=include_full),
            "economic_reserve": _serialize_array(shocked_res.economic_reserve, include_full=include_full),
            "index_s0": float(shocked_res.index_s0),
            "index_level_at_payment": _serialize_array(shocked_res.index_level_at_payment, include_full=include_full),
            "index_simple_return": _serialize_array(shocked_res.index_simple_return, include_full=include_full),
            "index_log_return": _serialize_array(shocked_res.index_log_return, include_full=include_full),
            "index_cumulative_return": _serialize_array(shocked_res.index_cumulative_return, include_full=include_full),
        },
        "tail_risk_mc": {
            "baseline": {
                "n_sims": int(getattr(baseline_mc, "n_sims", 0)),
                "premium_mean": float(getattr(baseline_mc, "premium_mean", float("nan"))),
                "premium_median": float(getattr(baseline_mc, "premium_median", float("nan"))),
                "premium_p05": float(getattr(baseline_mc, "premium_p05", float("nan"))),
                "premium_p95": float(getattr(baseline_mc, "premium_p95", float("nan"))),
            },
            "after": {
                "n_sims": int(getattr(shocked_mc, "n_sims", 0)),
                "premium_mean": float(getattr(shocked_mc, "premium_mean", float("nan"))),
                "premium_median": float(getattr(shocked_mc, "premium_median", float("nan"))),
                "premium_p05": float(getattr(shocked_mc, "premium_p05", float("nan"))),
                "premium_p95": float(getattr(shocked_mc, "premium_p95", float("nan"))),
            },
        },
    }
    if alm_base is not None:
        # Always include bucket time series for diagnostics completeness.
        out["alm_base"] = _alm_result_to_dict(alm_base, asm, include_buckets=True, include_full=include_full)
    else:
        out["alm_base"] = None
    if alm_after is not None:
        # Always include bucket time series for diagnostics completeness.
        out["alm_after"] = _alm_result_to_dict(alm_after, asm, include_buckets=True, include_full=include_full)
    else:
        out["alm_after"] = None
    return out


MortalityMode = Literal["synthetic", "qx_csv", "rp2014_mp2016", "us_ssa_2015_period"]
YieldMode = Literal["flat", "zero_csv", "par_bootstrap"]
ExpenseMode = Literal["csv", "manual"]

_SSA_2015_PERIOD_QX_CSV = """age,male_qx,female_qx
0,0.006383,0.005374
1,0.000453,0.000353
2,0.000282,0.000231
3,0.000230,0.000165
4,0.000169,0.000129
5,0.000155,0.000116
6,0.000145,0.000107
7,0.000135,0.000101
8,0.000120,0.000096
9,0.000105,0.000092
10,0.000094,0.000091
11,0.000099,0.000096
12,0.000134,0.000111
13,0.000207,0.000138
14,0.000309,0.000174
15,0.000419,0.000214
16,0.000530,0.000254
17,0.000655,0.000294
18,0.000791,0.000330
19,0.000934,0.000364
20,0.001085,0.000399
21,0.001228,0.000436
22,0.001339,0.000469
23,0.001403,0.000497
24,0.001433,0.000522
25,0.001451,0.000546
26,0.001475,0.000572
27,0.001502,0.000604
28,0.001538,0.000644
29,0.001581,0.000690
30,0.001626,0.000740
31,0.001669,0.000792
32,0.001712,0.000841
33,0.001755,0.000886
34,0.001800,0.000929
35,0.001855,0.000977
36,0.001920,0.001034
37,0.001988,0.001098
38,0.002060,0.001171
39,0.002141,0.001253
40,0.002240,0.001347
41,0.002362,0.001452
42,0.002509,0.001571
43,0.002684,0.001706
44,0.002890,0.001857
45,0.003121,0.002022
46,0.003386,0.002204
47,0.003707,0.002411
48,0.004091,0.002648
49,0.004531,0.002910
50,0.005013,0.003193
51,0.005524,0.003491
52,0.006059,0.003801
53,0.006611,0.004119
54,0.007187,0.004449
55,0.007800,0.004813
56,0.008456,0.005201
57,0.009144,0.005583
58,0.009865,0.005952
59,0.010622,0.006325
60,0.011458,0.006749
61,0.012350,0.007238
62,0.013235,0.007776
63,0.014097,0.008368
64,0.014979,0.009032
65,0.015967,0.009794
66,0.017109,0.010673
67,0.018392,0.011676
68,0.019836,0.012815
69,0.021465,0.014105
70,0.023351,0.015616
71,0.025482,0.017318
72,0.027794,0.019118
73,0.030282,0.020996
74,0.033022,0.023033
75,0.036201,0.025413
76,0.039858,0.028197
77,0.043891,0.031313
78,0.048311,0.034782
79,0.053228,0.038689
80,0.058897,0.043258
81,0.065365,0.048490
82,0.072491,0.054223
83,0.080288,0.060446
84,0.088916,0.067338
85,0.098576,0.075133
86,0.109438,0.084033
87,0.121619,0.094177
88,0.135176,0.105633
89,0.150109,0.118407
90,0.166397,0.132476
91,0.183997,0.147801
92,0.202855,0.164331
93,0.222911,0.182012
94,0.244094,0.200783
95,0.265091,0.219758
96,0.285508,0.238630
97,0.304926,0.257065
98,0.322919,0.274706
99,0.339065,0.291189
100,0.356018,0.308660
101,0.373819,0.327180
102,0.392510,0.346810
103,0.412135,0.367619
104,0.432742,0.389676
105,0.454379,0.413057
106,0.477098,0.437840
107,0.500953,0.464111
108,0.526000,0.491957
109,0.552300,0.521475
110,0.579915,0.552763
111,0.608911,0.585929
112,0.639357,0.621085
113,0.671325,0.658350
114,0.704891,0.697851
115,0.740135,0.739722
116,0.777142,0.777142
117,0.815999,0.815999
118,0.856799,0.856799
119,0.899639,0.899639
"""

SECTION_LABELS: dict[str, str] = {
    "overview": "Overview",
    "run": "Pricing Run",
    "alm": "ALM",
    "what_if": "What-if Analysis",
    "excel_replicator": "Excel Replicator",
    "tests": "Unit Tests",
}
SECTION_ORDER: list[str] = [
    "overview",
    "run",
    "alm",
    "what_if",
    "excel_replicator",
    "tests",
]

def _dynamic_overview_features() -> list[str]:
    options = list(product_options_for_ui())
    available_products = ", ".join(product_label(p) for p in options) if options else "None"
    mc_products = [product_label(p) for p in options if get_product_capabilities(p).supports_monte_carlo]
    econ_products = [product_label(p) for p in options if get_product_capabilities(p).supports_economic_scenario]
    return [
        f"Supported product run types: {available_products}.",
        "Run-time pricing dispatch is centralized in the product registry adapters.",
        f"Economic scenario controls enabled for: {', '.join(econ_products) if econ_products else 'None'}.",
        f"Monte Carlo pricing enabled for: {', '.join(mc_products) if mc_products else 'None'}.",
        "Yield curve sources: flat rate, zero-curve CSV, or par-yield CSV bootstrapped to zeros.",
        "Mortality sources are product-scoped and configured by registry defaults/options.",
        "ALM tab supports Treasury ladder projection, reinvestment/disinvestment policy controls, and KPI output tied to the active pricing run.",
        "What-if analysis provides before/after/impact views across pricing and ALM dimensions.",
        "Excel replicator export includes parity-oriented workbook output with optional MC and ALM snapshots.",
        "Embedded unit-test dashboard is available from the Unit Tests section.",
    ]


def _seed_run_form_state_from_last_inputs() -> None:
    meta = st.session_state.get("pricing_meta") or {}
    saved_inputs = st.session_state.get("pricing_run_inputs") or {}
    product_default = str(
        st.session_state.get("pricing_product_type", meta.get("product_type", ProductType.SPIA.value))
    )
    try:
        default_product_type = ProductType(product_default)
    except ValueError:
        default_product_type = ProductType.SPIA

    def _nonblank_str(saved_key: str, fallback: str) -> str:
        raw = saved_inputs.get(saved_key, fallback)
        txt = str(raw) if raw is not None else ""
        return txt if txt.strip() else fallback

    defaults: dict[str, Any] = {
        "run_product_type": product_default,
        "run_issue_age": int(saved_inputs.get("issue_age", 65)),
        "run_sex": str(saved_inputs.get("sex", "male")),
        "run_term_monthly_premium": float(saved_inputs.get("term_monthly_premium", 150.0)),
        "run_y_mode": str(meta.get("yield_mode", "par_bootstrap")),
        "run_m_mode": str(
            meta.get("mortality_mode", get_product_default_mortality_mode(default_product_type))
        ),
        "run_expense_mode": str(meta.get("expense_mode", "csv")),
        "run_horizon_age": int(saved_inputs.get("horizon_age", 110)),
        "run_valuation_year": int(saved_inputs.get("valuation_year", 2025)),
        "run_spread": float(saved_inputs.get("spread", 0.0)),
        "run_use_index": bool(saved_inputs.get("use_index", True)),
        "run_index_csv": str(saved_inputs.get("index_scenario_csv_path") or sp.DEFAULT_SP500_SCENARIO_CSV),
        "run_expense_inflation_pct": float(saved_inputs.get("expense_annual_inflation", 0.025) * 100.0),
        "run_mc_enable": bool(saved_inputs.get("mc_enabled", True)),
        "run_mc_n_sims": int(saved_inputs.get("mc_n_sims", 100)),
        "run_mc_seed": int(saved_inputs.get("mc_seed", 42)),
        "run_mc_drift_pct": float(saved_inputs.get("mc_annual_drift", 0.06) * 100.0),
        "run_mc_vol_pct": float(saved_inputs.get("mc_annual_vol", 0.15) * 100.0),
        "run_mc_s0": float(saved_inputs.get("mc_s0", 100.0)),
        "run_qx_csv": _nonblank_str("mortality_qx_csv", sp.DEFAULT_MORTALITY_QX_CSV),
        "run_rp_xlsx": _nonblank_str("mortality_rp_xlsx", sp.DEFAULT_RP2014_XLSX),
        "run_rp_out": _nonblank_str("mortality_rp_out_csv", sp.DEFAULT_RP2014_MALE_HEALTHY_QX_CSV),
        "run_mp_xlsx": _nonblank_str("mortality_mp_xlsx", sp.DEFAULT_MP2016_XLSX),
        "run_mp_out": _nonblank_str("mortality_mp_out_csv", sp.DEFAULT_MP2016_MALE_IMPROVEMENT_CSV),
    }
    for k, v in defaults.items():
        st.session_state.setdefault(k, v)


def _normalize_run_state_for_selected_product(
    state: MutableMapping[str, Any],
    *,
    selected_product: ProductType,
    switched_product: bool,
) -> None:
    """Normalize run-form state so UI values remain valid across product switches and reruns."""
    capabilities = get_product_capabilities(selected_product)

    # Keep enumerated controls valid for current product.
    mortality_options = list(get_product_mortality_mode_options(selected_product))
    default_mortality_mode = get_product_default_mortality_mode(selected_product)
    current_m_mode = str(state.get("run_m_mode", ""))
    if switched_product and selected_product == ProductType.SPIA:
        # SPIA should always land on RP+MP defaults when switching back from other products.
        state["run_m_mode"] = default_mortality_mode
    elif current_m_mode not in mortality_options:
        state["run_m_mode"] = default_mortality_mode

    y_mode = str(state.get("run_y_mode", "par_bootstrap"))
    if y_mode not in ("flat", "zero_csv", "par_bootstrap"):
        state["run_y_mode"] = "par_bootstrap"
    expense_mode = str(state.get("run_expense_mode", "csv"))
    if expense_mode not in ("csv", "manual"):
        state["run_expense_mode"] = "csv"
    if str(state.get("run_sex", "male")) not in ("male", "female"):
        state["run_sex"] = "male"

    # Keep path-like inputs nonblank; blank state often appears after product switching.
    if not str(state.get("run_index_csv", "")).strip():
        state["run_index_csv"] = sp.DEFAULT_SP500_SCENARIO_CSV
    if not str(state.get("run_qx_csv", "")).strip():
        state["run_qx_csv"] = sp.DEFAULT_MORTALITY_QX_CSV
    if not str(state.get("run_rp_xlsx", "")).strip():
        state["run_rp_xlsx"] = sp.DEFAULT_RP2014_XLSX
    if not str(state.get("run_rp_out", "")).strip():
        state["run_rp_out"] = sp.DEFAULT_RP2014_MALE_HEALTHY_QX_CSV
    if not str(state.get("run_mp_xlsx", "")).strip():
        state["run_mp_xlsx"] = sp.DEFAULT_MP2016_XLSX
    if not str(state.get("run_mp_out", "")).strip():
        state["run_mp_out"] = sp.DEFAULT_MP2016_MALE_IMPROVEMENT_CSV

    # Product capabilities govern whether these toggles should remain enabled.
    if not capabilities.supports_economic_scenario:
        state["run_use_index"] = False
    if not capabilities.supports_monte_carlo:
        state["run_mc_enable"] = False


def _build_yield_curve(
    mode: YieldMode,
    *,
    flat_rate: float,
    zero_csv: str,
    par_csv: str,
    coupon_freq: int,
) -> sp.YieldCurve:
    if mode == "flat":
        return sp.YieldCurve.from_flat_rate(float(flat_rate))
    if mode == "zero_csv":
        return sp.YieldCurve.load_zero_curve_csv(str(_resolve_path(zero_csv)))
    return sp.YieldCurve.load_par_yield_csv_and_bootstrap(
        str(_resolve_path(par_csv)),
        coupon_freq=int(coupon_freq),
    )


def _build_mortality(
    mode: MortalityMode,
    *,
    product_type: ProductType,
    sex: Literal["male", "female"],
    qx_csv: str,
    rp_xlsx: str,
    rp_out_csv: str,
    mp_xlsx: str,
    mp_out_csv: str,
) -> tuple[sp.MortalityTableQx | sp.MortalityTableRP2014MP2016, bool]:
    """
    Returns (mortality, needs_valuation_year).
    """
    if mode == "us_ssa_2015_period":
        if product_type != ProductType.TERM_LIFE:
            raise ValueError("US SSA 2015 period mortality is currently scoped to Term Life.")
        raw = pd.read_csv(io.StringIO(_SSA_2015_PERIOD_QX_CSV))
        qx_col = "male_qx" if sex == "male" else "female_qx"
        return sp.MortalityTableQx(raw["age"].to_numpy(dtype=int), raw[qx_col].to_numpy(dtype=float)), False
    if mode == "synthetic":
        return _minimal_mortality(), False
    if mode == "qx_csv":
        return sp.MortalityTableQx.load_qx_csv(str(_resolve_path(qx_csv))), False
    base_qx = sp.ensure_rp2014_male_healthy_annuitant_qx_csv(
        rp2014_xlsx_path=str(_resolve_path(rp_xlsx)),
        out_csv_path=str(_resolve_path(rp_out_csv)),
    )
    mp_ages, mp_years, mp_i = sp.ensure_mp2016_male_improvement_csv(
        mp2016_xlsx_path=str(_resolve_path(mp_xlsx)),
        out_csv_path=str(_resolve_path(mp_out_csv)),
    )
    mortality = sp.MortalityTableRP2014MP2016(
        base_qx_2014=base_qx,
        mp2016_ages=mp_ages,
        mp2016_years=mp_years,
        mp2016_i_matrix=mp_i,
        base_year=2014,
    )
    return mortality, True


def _render_overview() -> None:
    st.header("Model overview")
    st.markdown(
        "This workspace runs the pricing and projection engine with product adapters, "
        "scenario analysis, and Excel parity checks."
    )
    st.caption(
        "Overview content is generated from the product registry and shared section metadata "
        "to reduce documentation drift after model updates."
    )

    st.subheader("Current feature set")
    for i, feat in enumerate(_dynamic_overview_features(), start=1):
        st.markdown(f"{i}. {feat}")

    st.subheader("Workspace sections")
    section_labels = [SECTION_LABELS[k] for k in SECTION_ORDER if k != "overview"]
    st.markdown("Use the sidebar to navigate: " + " | ".join(f"**{name}**" for name in section_labels) + ".")


def _result_dataframe(res: sp.SPIAProjectionResult, contract: sp.SPIAContract) -> pd.DataFrame:
    expected_payment_pv = res.expected_benefit_cashflows * res.discount_factors
    cumulative_pv = np.cumsum(expected_payment_pv)
    return pd.DataFrame(
        {
            "month": res.months,
            "time_years": res.times_years,
            "age_at_payment": res.ages_at_payment,
            "survival": res.survival_to_payment,
            "discount_factor": res.discount_factors,
            "index_level": res.index_level_at_payment,
            "index_simple_return": res.index_simple_return,
            "index_log_return": res.index_log_return,
            "index_cumulative_return": res.index_cumulative_return,
            "benefit_nominal": res.benefit_nominal_scheduled,
            "expense_nominal": res.expense_nominal_scheduled,
            "expected_benefit": res.expected_benefit_cashflows,
            "expected_expense": res.expected_expense_cashflows,
            "expected_total": res.expected_total_cashflows,
            "expected_payment_pv": expected_payment_pv,
            "cumulative_benefit_pv": cumulative_pv,
        }
    )


def _render_pricing_run_charts(
    res: sp.SPIAProjectionResult, contract: sp.SPIAContract, expenses: sp.ExpenseAssumptions | None
) -> None:
    expected_payment_pv = res.expected_benefit_cashflows * res.discount_factors
    ages_r = contract.issue_age + res.reserve_times_years

    st.subheader("Run charts")
    st.markdown("**PV benefits**")
    st.line_chart(pd.DataFrame({"age": res.ages_at_payment, "pv_benefits": np.rint(expected_payment_pv)}).set_index("age"))

    st.markdown("**Economic reserve** (benefit + monthly expense, PV roll-forward)")
    st.line_chart(pd.DataFrame({"age": ages_r, "reserve": np.rint(res.economic_reserve)}).set_index("age"))

    if isinstance(expenses, sp.ExpenseAssumptions):
        _render_profit_decomposition_chart(res, contract, expenses)
    else:
        st.warning("Profit decomposition unavailable: pricing expense assumptions were not found in session state.")


def _render_profit_decomposition_chart(
    res: sp.SPIAProjectionResult, contract: sp.SPIAContract, expenses: sp.ExpenseAssumptions
) -> None:
    st.subheader("Profit decomposition waterfall")

    b_month = float(contract.benefit_annual) / 12.0
    n_months = int(res.months.size)

    level_benefit_certain_undisc = float(b_month * n_months)
    level_benefit_mort_undisc = float(np.sum(b_month * res.survival_to_payment))
    level_benefit_mort_disc = float(np.sum(b_month * res.survival_to_payment * res.discount_factors))

    mortality_effect = level_benefit_mort_undisc - level_benefit_certain_undisc
    discounting_effect = level_benefit_mort_disc - level_benefit_mort_undisc
    indexation_option_cost = float(res.pv_benefit - level_benefit_mort_disc)
    expense_component = float(expenses.policy_expense_dollars) + float(res.pv_monthly_expenses)
    margin_component = float(
        res.single_premium - (res.pv_benefit + res.pv_monthly_expenses + float(expenses.policy_expense_dollars))
    )

    rows = [
        ("Undiscounted level benefits (certain life)", level_benefit_certain_undisc, True),
        ("Mortality effect", mortality_effect, False),
        ("Discounting effect", discounting_effect, False),
        ("Indexation option cost", indexation_option_cost, False),
        ("Expenses (issue + monthly PV)", expense_component, False),
        ("Margin / premium load", margin_component, False),
        ("Single premium", float(res.single_premium), True),
    ]

    wf_rows = []
    running = 0.0
    for label, val, is_total in rows:
        if is_total:
            wf_rows.append({"Step": label, "delta": 0.0, "base": 0.0, "top": val, "is_total": True})
            running = val
            continue
        base = running if val >= 0.0 else running + val
        wf_rows.append({"Step": label, "delta": val, "base": base, "top": running + val, "is_total": False})
        running += val

    wf = pd.DataFrame(wf_rows)
    st.bar_chart(
        _round_for_visuals(wf.set_index("Step")[["base", "delta"]]),
        stack=True,
        color=["rgba(0,0,0,0)", "#1f77b4"],
    )

    table = pd.DataFrame(
        [
            {"Component": label, "Amount ($)": val}
            for label, val, _ in rows
            if label != "Single premium"
        ]
    )
    table_display = _round_for_visuals(table)
    st.dataframe(
        table_display,
        use_container_width=True,
        hide_index=True,
        column_config=_number_cols_no_decimals(table_display),
    )
    st.caption(
        "Interpretation: start from a high baseline that assumes level benefits are paid with certainty and "
        "without discounting. Mortality and discounting usually reduce that baseline because some future payments "
        "are not expected to occur and all future cashflows are worth less today. Indexation adds value back when "
        "projected indexed benefits exceed level benefits, while expenses and premium load/margin increase required "
        "premium further. The final total equals the modeled single premium."
    )


def _shock_yield_curve(curve: sp.YieldCurve, rate_shift_bps: float) -> sp.YieldCurve:
    shift = float(rate_shift_bps) / 10000.0
    return sp.YieldCurve(
        maturities_years=np.asarray(curve.maturities_years, dtype=float).copy(),
        zero_rates=np.asarray(curve.zero_rates, dtype=float).copy() + shift,
    )


def _key_rate_bump_curve(
    curve: sp.YieldCurve,
    *,
    key_tenor_years: float,
    key_tenors_years: np.ndarray,
    bump_bps: float = 1.0,
) -> sp.YieldCurve:
    return sp.yield_curve_key_rate_bump(
        curve,
        key_tenor_years=key_tenor_years,
        key_tenors_years=key_tenors_years,
        bump_bps=bump_bps,
    )


def _shock_mortality(
    mortality: sp.MortalityTableQx | sp.MortalityTableRP2014MP2016,
    longevity_improvement_pct: float,
) -> sp.MortalityTableQx | sp.MortalityTableRP2014MP2016:
    # Positive longevity improvement means lower qx.
    factor = max(0.01, 1.0 - float(longevity_improvement_pct) / 100.0)
    if isinstance(mortality, sp.MortalityTableQx):
        return sp.MortalityTableQx(
            ages=np.asarray(mortality.ages, dtype=int).copy(),
            qx=np.clip(np.asarray(mortality.qx, dtype=float) * factor, 0.0, 0.999999),
        )
    shocked_base = sp.MortalityTableQx(
        ages=np.asarray(mortality.base_qx_2014.ages, dtype=int).copy(),
        qx=np.clip(np.asarray(mortality.base_qx_2014.qx, dtype=float) * factor, 0.0, 0.999999),
    )
    return sp.MortalityTableRP2014MP2016(
        base_qx_2014=shocked_base,
        mp2016_ages=np.asarray(mortality.mp2016_ages, dtype=int).copy(),
        mp2016_years=np.asarray(mortality.mp2016_years, dtype=int).copy(),
        mp2016_i_matrix=np.asarray(mortality.mp2016_i_matrix, dtype=float).copy(),
        base_year=int(mortality.base_year),
    )


def _equity_regime_params(regime: str) -> tuple[float, float]:
    mapping = {
        "defensive": (0.03, 0.10),
        "base": (0.06, 0.15),
        "bullish": (0.09, 0.20),
        "stressed": (-0.02, 0.35),
    }
    return mapping.get(regime, mapping["base"])


def _deterministic_index_levels_from_regime(
    *, s0: float, annual_drift: float, n_months: int
) -> np.ndarray:
    dt = 1.0 / 12.0
    months = np.arange(1, n_months + 1, dtype=float)
    return float(s0) * np.exp(float(annual_drift) * months * dt)


def _render_impact_metric(label: str, before_val: float, after_val: float, money: bool = True) -> None:
    delta = float(after_val - before_val)
    if money:
        st.metric(label, f"${after_val:,.0f}", delta=f"${delta:,.0f}")
    else:
        st.metric(label, f"{after_val:,.0f}", delta=f"{delta:+,.0f}")


def _mc_cache_get_or_compute(
    key: tuple[object, ...],
    *,
    contract: sp.SPIAContract,
    yield_curve: sp.YieldCurve,
    mortality: sp.MortalityTableQx | sp.MortalityTableRP2014MP2016,
    horizon_age: int,
    spread: float,
    valuation_year: int | None,
    expenses: sp.ExpenseAssumptions,
    expense_annual_inflation: float,
    n_sims: int,
    annual_drift: float,
    annual_vol: float,
    seed: int,
    s0: float,
) -> sp.SPIAMonteCarloResult:
    cache = st.session_state.setdefault("what_if_mc_cache", {})
    if key in cache:
        return cache[key]
    out = sp.price_spia_single_premium_monte_carlo(
        contract=contract,
        yield_curve=yield_curve,
        mortality=mortality,
        horizon_age=horizon_age,
        spread=spread,
        valuation_year=valuation_year,
        expenses=expenses,
        expense_annual_inflation=expense_annual_inflation,
        n_sims=n_sims,
        annual_drift=annual_drift,
        annual_vol=annual_vol,
        seed=seed,
        s0=s0,
    )
    cache[key] = out
    return out


def _render_what_if_studio() -> None:
    st.header("What-if Analysis")
    st.caption("Live scenario shocks relative to the latest baseline run in Pricing Run.")

    base_res = st.session_state.get("pricing_res")
    base_contract = st.session_state.get("pricing_contract")
    ctx = st.session_state.get("pricing_excel_context") or {}
    base_curve = ctx.get("yield_curve")
    base_mort = ctx.get("mortality")
    base_expenses = ctx.get("expenses")

    if (
        base_res is None
        or base_contract is None
        or not isinstance(base_curve, sp.YieldCurve)
        or not isinstance(base_expenses, sp.ExpenseAssumptions)
        or not isinstance(base_mort, (sp.MortalityTableQx, sp.MortalityTableRP2014MP2016))
    ):
        st.info("Run pricing first in Pricing Run to set a baseline for What-if analysis.")
        return

    c1, c2, c3 = st.columns(3)
    with c1:
        rate_shift_bps = st.slider("Rates shift (bps)", min_value=-300, max_value=300, value=0, step=5)
        spread_shift_bps = st.slider("Credit spread shift (bps)", min_value=-300, max_value=300, value=0, step=5)
    with c2:
        inflation_shift_pct = st.slider("Expense inflation shift (%)", min_value=-5.0, max_value=10.0, value=0.0, step=0.1)
        longevity_improvement_pct = st.slider(
            "Longevity improvement shock (%)",
            min_value=-20.0,
            max_value=20.0,
            value=0.0,
            step=0.5,
            help="Positive values reduce mortality rates (longer lives).",
        )
    with c3:
        expense_ratio_mult = st.slider("Expense ratio multiplier", min_value=0.50, max_value=2.00, value=1.00, step=0.05)
        equity_regime = st.selectbox(
            "Equity regime",
            options=["defensive", "base", "bullish", "stressed"],
            index=1,
            format_func=lambda x: x.capitalize(),
        )
        mc_sims = st.slider("Tail-risk MC simulations", min_value=200, max_value=5000, value=800, step=200)

    st.markdown("**ALM add-on shocks**")
    st.caption(
        "These apply on top of the main What-if scenario. Asset curve shocks tilt **mark-to-market** on the Treasury ladder; "
        "liability stress scales **After** SPIA outflows in the ALM engine. "
        "Uses assumptions from the **ALM** tab if you ran them there; otherwise built-in defaults."
    )
    wa1, wa2, wa3, wa4 = st.columns(4)
    with wa1:
        alm_asset_parallel_bps = st.slider(
            "Asset earned-rate parallel shift (bps)",
            min_value=-200,
            max_value=200,
            value=0,
            step=5,
            help="Extra parallel shift on the **After** zero curve for **asset** discounting only.",
        )
    with wa2:
        alm_twist_short_bps = st.slider("Twist: short-end add-on (bps)", -75, 75, 0, 5)
    with wa3:
        alm_twist_long_bps = st.slider("Twist: long-end add-on (bps)", -75, 75, 0, 5)
    with wa4:
        alm_liability_cf_pct = st.slider(
            "Liability outflow stress (%)",
            -40.0,
            40.0,
            0.0,
            0.5,
            help="Scales **After** SPIA cash outflows in the ALM projection (stress on liquidity / disinvestment).",
        )

    alm_whatif_base: sp.ALMResult | None = None
    alm_whatif_after: sp.ALMResult | None = None
    asm_whatif_used: sp.ALMAssumptions | None = None
    try:
        horizon_age = int(ctx.get("horizon_age", 110))
        base_spread = float(ctx.get("spread", 0.0))
        valuation_year = ctx.get("valuation_year")
        base_infl = float(base_res.expense_annual_inflation)
        s0 = float(base_res.index_s0)
        n_months = int(base_res.months.size)

        shocked_curve = _shock_yield_curve(base_curve, float(rate_shift_bps))
        shocked_mort = _shock_mortality(base_mort, float(longevity_improvement_pct))
        shocked_expenses = sp.ExpenseAssumptions(
            policy_expense_dollars=float(base_expenses.policy_expense_dollars) * float(expense_ratio_mult),
            premium_expense_rate=min(0.99, float(base_expenses.premium_expense_rate) * float(expense_ratio_mult)),
            monthly_expense_dollars=float(base_expenses.monthly_expense_dollars) * float(expense_ratio_mult),
        )
        shocked_infl = max(-0.99, base_infl + float(inflation_shift_pct) / 100.0)
        shocked_spread = base_spread + float(spread_shift_bps) / 10000.0

        # Equity regime controls how the (deterministic) index levels used for "After" evolve.
        #
        # Key requirement: when What-if dials are at identity (all 0 / multipliers at 1) and
        # equity_regime == "base", "After" must reproduce the Pricing Run deterministic result.
        # We do that by applying a regime-specific multiplicative tilt to the Pricing Run's
        # actual baseline index_level_at_payment.
        drift_map, vol_map = _equity_regime_params(equity_regime)
        drift_base_map, vol_base_map = _equity_regime_params("base")

        base_is_identity = (
            equity_regime == "base"
            and abs(float(rate_shift_bps)) < 1e-12
            and abs(float(spread_shift_bps)) < 1e-12
            and abs(float(inflation_shift_pct)) < 1e-12
            and abs(float(longevity_improvement_pct)) < 1e-9
            and abs(float(expense_ratio_mult) - 1.0) < 1e-9
        )

        # Monte Carlo drift/vol are anchored to the Pricing Run's MC parameters so that
        # equity_regime=="base" gives an identity for tail-risk stats too.
        base_mc_params = st.session_state.get("pricing_mc_params") or {}
        base_drift = float(base_mc_params.get("annual_drift", 0.06))
        base_vol = float(base_mc_params.get("annual_vol", 0.15))

        if base_is_identity:
            idx_levels = np.asarray(base_res.index_level_at_payment, dtype=float)
        else:
            idx_regime_det = _deterministic_index_levels_from_regime(
                s0=s0, annual_drift=drift_map, n_months=n_months
            )
            idx_base_det = _deterministic_index_levels_from_regime(
                s0=s0, annual_drift=drift_base_map, n_months=n_months
            )
            idx_base_det = np.asarray(idx_base_det, dtype=float)
            idx_regime_det = np.asarray(idx_regime_det, dtype=float)
            scale = idx_regime_det / idx_base_det
            idx_levels = np.asarray(base_res.index_level_at_payment, dtype=float) * scale

        if equity_regime == "base":
            drift_mc = base_drift
            vol_mc = base_vol
        else:
            # Scale regime drifts/vols relative to the regime mapping's "base" so the meaning
            # of defensive/bullish/stressed stays consistent even if the Pricing Run used different MC inputs.
            drift_mc = base_drift * (drift_map / drift_base_map) if abs(drift_base_map) > 1e-15 else base_drift
            vol_mc = base_vol * (vol_map / vol_base_map) if abs(vol_base_map) > 1e-15 else base_vol

        shocked_res = sp.price_spia_single_premium(
            contract=base_contract,
            yield_curve=shocked_curve,
            mortality=shocked_mort,
            horizon_age=horizon_age,
            spread=shocked_spread,
            valuation_year=int(valuation_year) if valuation_year is not None else None,
            expenses=shocked_expenses,
            index_s0=s0,
            index_levels_payment=idx_levels,
            expense_annual_inflation=shocked_infl,
        )
        vy = int(valuation_year) if valuation_year is not None else None
        baseline_key = (
            "baseline",
            int(mc_sims),
            int(horizon_age),
            float(base_spread),
            float(base_infl),
            float(base_drift),
            float(base_vol),
            float(s0),
            int(base_contract.issue_age),
            float(base_contract.benefit_annual),
        )
        baseline_mc = _mc_cache_get_or_compute(
            baseline_key,
            contract=base_contract,
            yield_curve=base_curve,
            mortality=base_mort,
            horizon_age=horizon_age,
            spread=base_spread,
            valuation_year=vy,
            expenses=base_expenses,
            expense_annual_inflation=base_infl,
            n_sims=int(mc_sims),
            annual_drift=float(base_drift),
            annual_vol=float(base_vol),
            seed=42,
            s0=float(s0),
        )
        shocked_key = (
            "shocked",
            int(mc_sims),
            int(horizon_age),
            float(shocked_spread),
            float(shocked_infl),
            float(drift_mc),
            float(vol_mc),
            float(s0),
            float(rate_shift_bps),
            float(spread_shift_bps),
            float(longevity_improvement_pct),
            float(expense_ratio_mult),
            str(equity_regime),
            int(base_contract.issue_age),
            float(base_contract.benefit_annual),
        )
        shocked_mc = _mc_cache_get_or_compute(
            shocked_key,
            contract=base_contract,
            yield_curve=shocked_curve,
            mortality=shocked_mort,
            horizon_age=horizon_age,
            spread=shocked_spread,
            valuation_year=vy,
            expenses=shocked_expenses,
            expense_annual_inflation=shocked_infl,
            n_sims=int(mc_sims),
            annual_drift=float(drift_mc),
            annual_vol=float(vol_mc),
            seed=42,
            s0=float(s0),
        )

        try:
            asm_wf = st.session_state.get("alm_last_assumptions")
            if not isinstance(asm_wf, sp.ALMAssumptions):
                asm_wf = st.session_state.get("alm_current_assumptions")
            if not isinstance(asm_wf, sp.ALMAssumptions):
                asm_wf = sp.ALMAssumptions(
                    allocation=sp.alm_default_allocation_spec(),
                    rebalance_band=0.05,
                    rebalance_frequency_months=1,
                    reinvest_rule="pro_rata",
                    disinvest_rule="shortest_first",
                    rebalance_policy="liquidity_only",
                    borrowing_policy="borrow_after_assets_insufficient",
                    borrowing_rate_mode="scenario_linked",
                    borrowing_rate_tenor_years=1.0,
                    borrowing_spread_annual=0.01,
                    borrowing_rate_annual=0.05,
                    liquidity_near_liquid_years=0.25,
                )
            asm_whatif_used = asm_wf
            aum_wf = st.session_state.get("alm_last_initial_asset_market_value")
            if aum_wf is None:
                aum_wf = st.session_state.get("alm_current_initial_asset_market_value")
            aum_wf_use = float(aum_wf) if isinstance(aum_wf, (int, float, np.floating)) else float(base_res.single_premium)
            alm_whatif_base = sp.run_alm_projection(
                pricing=base_res,
                yield_curve=base_curve,
                spread=base_spread,
                assumptions=asm_wf,
                initial_asset_market_value=aum_wf_use,
            )
            yc_alm_asset = sp.yield_curve_twist_linear_bps(
                sp.yield_curve_parallel_bps(shocked_curve, float(alm_asset_parallel_bps)),
                bps_short=float(alm_twist_short_bps),
                bps_long=float(alm_twist_long_bps),
            )
            cf_alm = np.asarray(shocked_res.expected_total_cashflows, dtype=float) * (
                1.0 + float(alm_liability_cf_pct) / 100.0
            )
            alm_whatif_after = sp.run_alm_projection(
                pricing=shocked_res,
                yield_curve=shocked_curve,
                spread=shocked_spread,
                assumptions=asm_wf,
                initial_asset_market_value=aum_wf_use,
                asset_curve=yc_alm_asset,
                liability_cashflows=cf_alm,
            )
        except Exception as alm_ex:
            alm_whatif_base = None
            alm_whatif_after = None
            asm_whatif_used = None
            st.warning(f"ALM what-if layer skipped: {alm_ex!r}")
    except Exception as ex:
        st.error(f"What-if scenario failed: {ex!r}")
        return

    st.session_state["whatif_last_params"] = {
        "rates_shift_bps": float(rate_shift_bps),
        "spread_shift_bps": float(spread_shift_bps),
        "inflation_shift_pct": float(inflation_shift_pct),
        "longevity_improvement_pct": float(longevity_improvement_pct),
        "expense_ratio_mult": float(expense_ratio_mult),
        "equity_regime": str(equity_regime),
        "mc_sims": int(mc_sims),
        "alm_asset_parallel_bps": float(alm_asset_parallel_bps),
        "alm_twist_short_bps": float(alm_twist_short_bps),
        "alm_twist_long_bps": float(alm_twist_long_bps),
        "alm_liability_cf_pct": float(alm_liability_cf_pct),
    }
    st.session_state["whatif_last_base_res"] = base_res
    st.session_state["whatif_last_shocked_res"] = shocked_res
    st.session_state["whatif_last_baseline_mc"] = baseline_mc
    st.session_state["whatif_last_shocked_mc"] = shocked_mc
    st.session_state["whatif_last_shocked_curve"] = shocked_curve
    st.session_state["whatif_last_shocked_mortality"] = shocked_mort
    st.session_state["whatif_last_alm_base"] = alm_whatif_base
    st.session_state["whatif_last_alm_after"] = alm_whatif_after
    st.session_state["whatif_last_alm_assumptions"] = asm_whatif_used
    st.session_state["whatif_last_pricing_run_id"] = st.session_state.get("pricing_run_id")
    _invalidate_diagnostics_export()

    st.subheader("Before vs after vs impact")
    m1, m2, m3, m4 = st.columns(4)
    with m1:
        _render_impact_metric("Single premium", float(base_res.single_premium), float(shocked_res.single_premium), money=True)
    with m2:
        base_margin = float(base_res.single_premium - (base_res.pv_benefit + base_res.pv_monthly_expenses))
        shocked_margin = float(shocked_res.single_premium - (shocked_res.pv_benefit + shocked_res.pv_monthly_expenses))
        _render_impact_metric("Margin", base_margin, shocked_margin, money=True)
    with m3:
        _render_impact_metric("Reserve at issue", float(base_res.economic_reserve[0]), float(shocked_res.economic_reserve[0]), money=True)
    with m4:
        _render_impact_metric(
            "Tail risk (P95 premium)",
            float(baseline_mc.premium_p95),
            float(shocked_mc.premium_p95),
            money=True,
        )

    compare_df = pd.DataFrame(
        {
            "Metric": ["Single premium", "Margin", "Reserve at issue", "Tail risk (P95 premium)"],
            "Before": [
                float(base_res.single_premium),
                float(base_res.single_premium - (base_res.pv_benefit + base_res.pv_monthly_expenses)),
                float(base_res.economic_reserve[0]),
                float(baseline_mc.premium_p95),
            ],
            "After": [
                float(shocked_res.single_premium),
                float(shocked_res.single_premium - (shocked_res.pv_benefit + shocked_res.pv_monthly_expenses)),
                float(shocked_res.economic_reserve[0]),
                float(shocked_mc.premium_p95),
            ],
        }
    )
    compare_df["Impact"] = compare_df["After"] - compare_df["Before"]
    compare_display = _round_for_visuals(compare_df)
    st.dataframe(
        compare_display,
        use_container_width=True,
        hide_index=True,
        column_config=_number_cols_no_decimals(compare_display),
    )

    st.markdown("**Reserve path impact**")
    reserve_df = pd.DataFrame(
        {
            "age": base_contract.issue_age + base_res.reserve_times_years,
            "Before reserve": base_res.economic_reserve,
            "After reserve": shocked_res.economic_reserve,
            "Impact": shocked_res.economic_reserve - base_res.economic_reserve,
        }
    ).set_index("age")
    reserve_display = _round_for_visuals(reserve_df)
    # Clean up x-axis labels and monetary formatting.
    reserve_display.index = np.round(reserve_display.index.values.astype(float), 2)
    for col in ["Before reserve", "After reserve", "Impact"]:
        reserve_display[col] = reserve_display[col].astype(int)
    st.line_chart(reserve_display[["Before reserve", "After reserve"]])
    st.bar_chart(reserve_display[["Impact"]])

    st.markdown("**Tail-risk distribution impact (single premium)**")
    c1, c2 = st.columns(2)
    with c1:
        counts_b, edges_b = np.histogram(baseline_mc.single_premium, bins=35)
        mids_b = 0.5 * (edges_b[:-1] + edges_b[1:])
        mids_b_disp = np.rint(mids_b).astype(int)
        bin_labels_b = [f"{int(v):,}" for v in mids_b_disp]
        df_b = pd.DataFrame({"bin": bin_labels_b, "count_before": counts_b.astype(int)}).set_index("bin")
        st.bar_chart(_round_for_visuals(df_b))
    with c2:
        counts_a, edges_a = np.histogram(shocked_mc.single_premium, bins=35)
        mids_a = 0.5 * (edges_a[:-1] + edges_a[1:])
        mids_a_disp = np.rint(mids_a).astype(int)
        bin_labels_a = [f"{int(v):,}" for v in mids_a_disp]
        df_a = pd.DataFrame({"bin": bin_labels_a, "count_after": counts_a.astype(int)}).set_index("bin")
        st.bar_chart(_round_for_visuals(df_a))

    if alm_whatif_base is not None and alm_whatif_after is not None:
        _alm_b = alm_whatif_base
        _alm_a = alm_whatif_after
        st.subheader("ALM KPI impact")
        st.caption(
            "Before = ALM on the **Pricing Run** baseline; After = ALM on the shocked liability pricing with **After** curve "
            "for liability PV, optional extra asset **earned-rate** shifts, and scaled outflows. "
            "**Liquidity buffer** = (cash + bonds within near-liquid residual maturity) divided by mean expected monthly "
            "outflow over the next 12 months."
        )

        def _alm_snap(r: sp.ALMResult) -> dict[str, float]:
            return {
                "fr_m1": float(r.funding_ratio[0]) if r.funding_ratio.size else float("nan"),
                "surp_m1": float(r.surplus[0]) if r.surplus.size else float("nan"),
                "liq_m1": float(r.liquidity_buffer_months[0]) if r.liquidity_buffer_months.size else float("nan"),
                "pv01_net": float(r.pv01_net),
                "dur_gap": float(r.duration_gap),
            }

        sb = _alm_snap(_alm_b)
        sa = _alm_snap(_alm_a)
        alm_cmp = pd.DataFrame(
            {
                "Metric": [
                    "Funding ratio (month-end 1)",
                    "Surplus ($)",
                    "Liquidity buffer (months)",
                    "PV01 net ($ per 1bp)",
                    "Duration gap (years)",
                ],
                "Before": [
                    sb["fr_m1"],
                    sb["surp_m1"],
                    sb["liq_m1"],
                    sb["pv01_net"],
                    sb["dur_gap"],
                ],
                "After": [
                    sa["fr_m1"],
                    sa["surp_m1"],
                    sa["liq_m1"],
                    sa["pv01_net"],
                    sa["dur_gap"],
                ],
            }
        )
        alm_cmp["Impact"] = alm_cmp["After"] - alm_cmp["Before"]
        alm_show = alm_cmp.copy()
        alm_show[["Before", "After", "Impact"]] = alm_show[["Before", "After", "Impact"]].round(4)
        money_rows = alm_show["Metric"].isin(["Surplus ($)", "PV01 net ($ per 1bp)"])
        if bool(money_rows.any()):
            alm_show.loc[money_rows, ["Before", "After", "Impact"]] = alm_show.loc[
                money_rows, ["Before", "After", "Impact"]
            ].round(0).astype(int)
        st.dataframe(alm_show, use_container_width=True, hide_index=True)

        age_alm = np.round((base_contract.issue_age + _alm_b.times_years).astype(float), 2)
        path_cmp = pd.DataFrame(
            {
                "Funding ratio (before)": _alm_b.funding_ratio,
                "Funding ratio (after)": _alm_a.funding_ratio,
            },
            index=age_alm,
        )
        st.markdown("**Funding ratio path**")
        st.line_chart(path_cmp)
        sur_cmp = pd.DataFrame(
            {"Surplus before": _alm_b.surplus, "Surplus after": _alm_a.surplus},
            index=age_alm,
        )
        st.markdown("**Surplus path**")
        sur_disp = _round_for_visuals(sur_cmp)
        sur_disp[["Surplus before", "Surplus after"]] = sur_disp[["Surplus before", "Surplus after"]].astype(int)
        st.line_chart(sur_disp)

        st.markdown("**PV assets and liabilities**")
        pv_cmp = pd.DataFrame(
            {
                "PV assets (before)": _alm_b.asset_market_value,
                "PV assets (after)": _alm_a.asset_market_value,
                "PV liabilities (before)": _alm_b.liability_pv,
                "PV liabilities (after)": _alm_a.liability_pv,
            },
            index=age_alm,
        )
        pv_disp = _round_for_visuals(pv_cmp)
        for c in [
            "PV assets (before)",
            "PV assets (after)",
            "PV liabilities (before)",
            "PV liabilities (after)",
        ]:
            pv_disp[c] = pv_disp[c].astype(int)
        st.line_chart(pv_disp)

        st.markdown("**ALM key rate duration (before vs after)**")
        try:
            asm_krd_wf = asm_whatif_used if isinstance(asm_whatif_used, sp.ALMAssumptions) else None
            if asm_krd_wf is not None:
                key_tenors = np.array(
                    [float(b.tenor_years) for b in asm_krd_wf.allocation.buckets if float(b.tenor_years) > 1e-12],
                    dtype=float,
                )
                if key_tenors.size > 0:
                    a0_wf = st.session_state.get("alm_last_initial_asset_market_value")
                    if not isinstance(a0_wf, (int, float, np.floating)):
                        a0_wf = st.session_state.get("alm_current_initial_asset_market_value")
                    a0 = float(a0_wf) if isinstance(a0_wf, (int, float, np.floating)) else float(base_res.single_premium)

                    def _compute_krd_set(
                        *,
                        curve_liab: sp.YieldCurve,
                        curve_asset: sp.YieldCurve,
                        spread_use: float,
                        cashflows_use: np.ndarray,
                        scenario_label: str,
                    ) -> list[dict[str, float | str]]:
                        w_use = np.asarray(asm_krd_wf.allocation.weights, dtype=float)
                        bond_tenors = np.array(
                            [float(b.tenor_years) for b in asm_krd_wf.allocation.buckets[1:]],
                            dtype=float,
                        )
                        df0_asset = curve_asset.discount_factors(bond_tenors, spread=spread_use)
                        target_mv_bonds = w_use[1:] * a0
                        bond_faces = np.where(df0_asset > 1e-15, target_mv_bonds / df0_asset, 0.0)
                        l0 = float(np.sum(cashflows_use * curve_liab.discount_factors(base_res.times_years, spread=spread_use)))
                        net0 = max(1e-9, a0 - l0)
                        out_rows: list[dict[str, float | str]] = []
                        for kt in key_tenors:
                            cl_b = _key_rate_bump_curve(
                                curve_liab,
                                key_tenor_years=float(kt),
                                key_tenors_years=key_tenors,
                                bump_bps=1.0,
                            )
                            ca_b = _key_rate_bump_curve(
                                curve_asset,
                                key_tenor_years=float(kt),
                                key_tenors_years=key_tenors,
                                bump_bps=1.0,
                            )
                            a_b = float(w_use[0] * a0 + np.sum(bond_faces * ca_b.discount_factors(bond_tenors, spread=spread_use)))
                            l_b = float(np.sum(cashflows_use * cl_b.discount_factors(base_res.times_years, spread=spread_use)))
                            out_rows.extend(
                                [
                                    {
                                        "Tenor": f"{kt:g}Y",
                                        "Tenor years": float(kt),
                                        "Scenario": scenario_label,
                                        "Series": "Assets KRD",
                                        "KRD": -((a_b - a0) / (max(1e-9, a0) * 1e-4)),
                                    },
                                    {
                                        "Tenor": f"{kt:g}Y",
                                        "Tenor years": float(kt),
                                        "Scenario": scenario_label,
                                        "Series": "Liabilities KRD",
                                        "KRD": -((l_b - l0) / (max(1e-9, l0) * 1e-4)),
                                    },
                                    {
                                        "Tenor": f"{kt:g}Y",
                                        "Tenor years": float(kt),
                                        "Scenario": scenario_label,
                                        "Series": "Surplus KRD",
                                        "KRD": -(((a_b - l_b) - (a0 - l0)) / (net0 * 1e-4)),
                                    },
                                ]
                            )
                        return out_rows

                    krd_rows = []
                    krd_rows.extend(
                        _compute_krd_set(
                            curve_liab=base_curve,
                            curve_asset=base_curve,
                            spread_use=float(base_spread),
                            cashflows_use=np.asarray(base_res.expected_total_cashflows, dtype=float),
                            scenario_label="Before",
                        )
                    )
                    krd_rows.extend(
                        _compute_krd_set(
                            curve_liab=shocked_curve,
                            curve_asset=yc_alm_asset,
                            spread_use=float(shocked_spread),
                            cashflows_use=np.asarray(cf_alm, dtype=float),
                            scenario_label="After",
                        )
                    )
                    krd_wf_df = pd.DataFrame(krd_rows).sort_values(["Series", "Tenor years", "Scenario"])
                    tenor_order = [f"{float(t):g}Y" for t in np.sort(np.unique(key_tenors))]
                    series_order = ["Assets KRD", "Liabilities KRD", "Surplus KRD"]
                    # Faceted specs often render with a narrow default plot width in Streamlit (squished left,
                    # empty space right). Use stacked panels with container width so plots fill the column.
                    enc_x = alt.X(
                        "Tenor:N",
                        sort=tenor_order,
                        title="Key tenor",
                        axis=alt.Axis(labelAngle=0, labelPadding=4),
                    )

                    def _wf_krd_panel(subtitle: str, df_sub: pd.DataFrame, *, show_legend: bool) -> alt.Chart:
                        color_enc = alt.Color(
                            "Scenario:N",
                            sort=["Before", "After"],
                            scale=alt.Scale(domain=["Before", "After"], range=["#4c78a8", "#f58518"]),
                            legend=alt.Legend(orient="top", direction="horizontal") if show_legend else None,
                        )
                        return (
                            alt.Chart(df_sub)
                            .mark_line(point=True, strokeWidth=2.5)
                            .encode(
                                x=enc_x,
                                y=alt.Y("KRD:Q", title="KRD (years)"),
                                color=color_enc,
                                tooltip=[
                                    alt.Tooltip("Series:N"),
                                    alt.Tooltip("Tenor:N"),
                                    alt.Tooltip("Scenario:N"),
                                    alt.Tooltip("KRD:Q", format=".4f"),
                                ],
                            )
                            .properties(width="container", height=115, title=subtitle)
                        )

                    panels = [
                        _wf_krd_panel(
                            s,
                            krd_wf_df[krd_wf_df["Series"] == s],
                            show_legend=(i == 0),
                        )
                        for i, s in enumerate(series_order)
                    ]
                    krd_wf_chart = (
                        alt.vconcat(*panels, spacing=8)
                        .resolve_scale(y="independent")
                        .configure_view(strokeWidth=0)
                    )
                    st.altair_chart(krd_wf_chart, use_container_width=True)
                    st.caption(
                        "Each panel compares Before vs After KRD at key tenors for one series, with independent y-scales to keep "
                        "the view readable when Surplus KRD magnitudes are much larger."
                    )
                else:
                    st.info("No positive tenors available for What-if ALM KRD chart.")
            else:
                st.info("What-if ALM KRD chart unavailable: ALM assumptions were not available.")
        except Exception as ex:
            st.info(f"What-if ALM KRD chart unavailable for current inputs: {ex!r}")

    st.caption(
        "Impact shown as After - Before. Tail risk uses the 95th percentile of simulated premiums under the selected equity regime."
    )


def _render_run_and_results() -> None:
    st.header("Pricing Run")
    _seed_run_form_state_from_last_inputs()
    st.markdown(
        """
        <style>
            .product-type-callout {
                border: 2px solid #1f77b4;
                border-radius: 10px;
                padding: 10px 12px;
                background: rgba(31, 119, 180, 0.08);
                margin-bottom: 10px;
            }
            .product-type-callout strong {
                font-size: 1.05rem;
            }
        </style>
        <div class="product-type-callout">
            <strong>Primary input: Product Type</strong><br/>
            This selection controls which pricing engine, assumptions, and downstream outputs are active for this run.
        </div>
        """,
        unsafe_allow_html=True,
    )
    product_options = list(product_options_for_ui())
    product_values = [p.value for p in product_options]
    if st.session_state.get("run_product_type") not in product_values and product_values:
        st.session_state["run_product_type"] = product_values[0]
    selected_product = st.selectbox(
        "Product type",
        options=product_values,
        format_func=lambda raw: product_label(ProductType(raw)),
        help="Run exactly one product per execution.",
        key="run_product_type",
    )
    selected_product = ProductType(selected_product)
    last_product_raw = st.session_state.get("_run_last_product_type")
    switched_product = last_product_raw is not None and str(last_product_raw) != selected_product.value
    _normalize_run_state_for_selected_product(
        st.session_state,
        selected_product=selected_product,
        switched_product=switched_product,
    )
    st.session_state["_run_last_product_type"] = selected_product.value
    product_ui_cfg = get_product_ui_config(selected_product)
    if product_ui_cfg.selected_info_message:
        st.info(product_ui_cfg.selected_info_message)

    with st.expander("Contract", expanded=True):
        c1, c2, c3 = st.columns(3)
        issue_age = c1.number_input("Issue age", min_value=0, max_value=120, step=1, key="run_issue_age")
        sex = c2.selectbox("Sex (metadata)", options=["male", "female"], key="run_sex")
        if selected_product == ProductType.TERM_LIFE:
            term_ui = get_term_contract_ui_config()
            st.session_state.setdefault("run_term_benefit_annual", float(term_ui.default_death_benefit))
            benefit_annual = c3.number_input(
                term_ui.death_benefit_label,
                min_value=1.0,
                step=10_000.0,
                key="run_term_benefit_annual",
            )
            t1, t2, t3 = st.columns(3)
            st.session_state.setdefault("run_term_length", term_ui.term_length_options[0])
            st.session_state.setdefault("run_term_premium_mode", term_ui.premium_mode_options[0])
            st.session_state.setdefault("run_term_benefit_timing", term_ui.benefit_timing_options[0])
            term_choice = t1.selectbox("Term length", options=list(term_ui.term_length_options), key="run_term_length")
            premium_mode_choice = t2.selectbox(
                "Premium mode", options=list(term_ui.premium_mode_options), key="run_term_premium_mode"
            )
            benefit_timing_choice = t3.selectbox(
                "Benefit timing", options=list(term_ui.benefit_timing_options), key="run_term_benefit_timing"
            )
            monthly_premium = st.number_input(
                "Monthly premium ($)",
                min_value=0.0,
                step=10.0,
                key="run_term_monthly_premium",
            )
        else:
            st.session_state.setdefault("run_spia_benefit_annual", 100_000.0)
            benefit_annual = c3.number_input("Annual benefit ($)", min_value=0.0, step=1_000.0, key="run_spia_benefit_annual")
            term_choice = "n/a"
            premium_mode_choice = "n/a"
            benefit_timing_choice = "n/a"
            monthly_premium = 0.0

    with st.expander("Yield curve", expanded=True):
        y_mode = st.radio(
            "Source",
            options=["flat", "zero_csv", "par_bootstrap"],
            format_func=lambda x: {
                "flat": "Flat zero rate",
                "zero_csv": "Zero curve CSV",
                "par_bootstrap": "Par yields CSV → bootstrap zeros",
            }[x],
            horizontal=True,
            key="run_y_mode",
        )
        st.session_state.setdefault("run_flat_rate", 0.04)
        st.session_state.setdefault("run_zero_csv", sp.DEFAULT_ZERO_CURVE_CSV)
        st.session_state.setdefault("run_par_csv", sp.DEFAULT_PAR_CURVE_CSV)
        st.session_state.setdefault("run_coupon_freq", 2)
        flat_rate = float(st.session_state.get("run_flat_rate", 0.04))
        zero_csv = str(st.session_state.get("run_zero_csv", sp.DEFAULT_ZERO_CURVE_CSV))
        par_csv = str(st.session_state.get("run_par_csv", sp.DEFAULT_PAR_CURVE_CSV))
        coupon_freq = int(st.session_state.get("run_coupon_freq", 2))
        if y_mode == "flat":
            flat_rate = st.number_input("Flat continuously compounded zero rate", format="%.4f", key="run_flat_rate")
        elif y_mode == "zero_csv":
            zero_csv = st.text_input("Zero curve CSV path", key="run_zero_csv")
        else:
            par_csv = st.text_input("Par yield CSV path", key="run_par_csv")
            coupon_freq = st.number_input("Coupon payments per year", min_value=1, step=1, key="run_coupon_freq")

    with st.expander("Mortality", expanded=True):
        mortality_options = list(get_product_mortality_mode_options(selected_product))
        if st.session_state.get("run_m_mode") not in mortality_options:
            st.session_state["run_m_mode"] = get_product_default_mortality_mode(selected_product)
        mortality_default_mode = get_product_default_mortality_mode(selected_product)
        mortality_default_index = (
            mortality_options.index(mortality_default_mode)
            if mortality_default_mode in mortality_options
            else 0
        )
        m_mode = st.radio(
            "Table",
            options=mortality_options,
            format_func=lambda x: get_mortality_mode_label(str(x)),
            horizontal=True,
            index=mortality_default_index,
            key="run_m_mode",
        )
        st.session_state.setdefault("run_qx_csv", sp.DEFAULT_MORTALITY_QX_CSV)
        st.session_state.setdefault("run_rp_xlsx", sp.DEFAULT_RP2014_XLSX)
        st.session_state.setdefault("run_rp_out", sp.DEFAULT_RP2014_MALE_HEALTHY_QX_CSV)
        st.session_state.setdefault("run_mp_xlsx", sp.DEFAULT_MP2016_XLSX)
        st.session_state.setdefault("run_mp_out", sp.DEFAULT_MP2016_MALE_IMPROVEMENT_CSV)
        qx_csv = str(st.session_state.get("run_qx_csv", sp.DEFAULT_MORTALITY_QX_CSV))
        rp_xlsx = str(st.session_state.get("run_rp_xlsx", sp.DEFAULT_RP2014_XLSX))
        rp_out = str(st.session_state.get("run_rp_out", sp.DEFAULT_RP2014_MALE_HEALTHY_QX_CSV))
        mp_xlsx = str(st.session_state.get("run_mp_xlsx", sp.DEFAULT_MP2016_XLSX))
        mp_out = str(st.session_state.get("run_mp_out", sp.DEFAULT_MP2016_MALE_IMPROVEMENT_CSV))
        if m_mode == "qx_csv":
            qx_csv = st.text_input("q_x CSV (columns age, qx)", key="run_qx_csv")
        elif m_mode == "rp2014_mp2016":
            st.caption("SOA workbooks are optional if matching CSV extracts already exist beside the xlsx paths.")
            if not str(st.session_state.get("run_rp_xlsx", "")).strip():
                st.session_state["run_rp_xlsx"] = sp.DEFAULT_RP2014_XLSX
            if not str(st.session_state.get("run_rp_out", "")).strip():
                st.session_state["run_rp_out"] = sp.DEFAULT_RP2014_MALE_HEALTHY_QX_CSV
            if not str(st.session_state.get("run_mp_xlsx", "")).strip():
                st.session_state["run_mp_xlsx"] = sp.DEFAULT_MP2016_XLSX
            if not str(st.session_state.get("run_mp_out", "")).strip():
                st.session_state["run_mp_out"] = sp.DEFAULT_MP2016_MALE_IMPROVEMENT_CSV
            rp_xlsx = st.text_input("RP-2014 xlsx", key="run_rp_xlsx")
            rp_out = st.text_input("RP-2014 healthy male qx cache CSV", key="run_rp_out")
            mp_xlsx = st.text_input("MP-2016 xlsx", key="run_mp_xlsx")
            mp_out = st.text_input("MP-2016 improvement cache CSV", key="run_mp_out")
        elif m_mode == "us_ssa_2015_period":
            st.caption("Source: SSA actuarial life table (US Social Security area population), period year 2015.")

    with st.expander("Expenses & valuation", expanded=True):
        expense_mode = st.radio(
            "Expenses",
            options=["csv", "manual"],
            format_func=lambda x: "Load from CSV" if x == "csv" else "Enter manually",
            horizontal=True,
            key="run_expense_mode",
        )
        st.session_state.setdefault("run_expenses_csv", sp.DEFAULT_EXPENSES_CSV)
        st.session_state.setdefault("run_policy_expense", 0.0)
        st.session_state.setdefault("run_premium_expense_pct", 0.0)
        st.session_state.setdefault("run_monthly_expense", 0.0)
        expenses_csv = str(st.session_state.get("run_expenses_csv", sp.DEFAULT_EXPENSES_CSV))
        pol = float(st.session_state.get("run_policy_expense", 0.0))
        prem_pct = float(st.session_state.get("run_premium_expense_pct", 0.0))
        monthly_ex = float(st.session_state.get("run_monthly_expense", 0.0))
        if expense_mode == "csv":
            expenses_csv = st.text_input("Expenses CSV path", key="run_expenses_csv")
        else:
            pol = float(st.number_input("Policy expense at issue ($)", key="run_policy_expense"))
            prem_pct = float(
                st.number_input(
                    "Premium expense (% of single premium)",
                    min_value=0.0,
                    max_value=99.99,
                    help="Enter 2 for 2%. Must stay below 100%.",
                    key="run_premium_expense_pct",
                )
            )
            monthly_ex = float(st.number_input("Monthly expense while alive ($)", key="run_monthly_expense"))
        valuation_year = st.number_input(
            "Valuation year (calendar)",
            min_value=1950,
            max_value=2100,
            help="Used for RP+MP calendar-year mortality; ignored for static/synthetic q_x.",
            key="run_valuation_year",
        )
        horizon_age = st.number_input("Horizon age (stop monthly grid)", min_value=1, max_value=130, key="run_horizon_age")
        spread = st.number_input("Credit spread added to zero rate", format="%.4f", key="run_spread")

    product_caps = get_product_capabilities(selected_product)
    can_use_economic_scenario = bool(product_caps.supports_economic_scenario)
    can_use_monte_carlo = bool(product_caps.supports_monte_carlo)

    use_index = False
    index_csv = sp.DEFAULT_SP500_SCENARIO_CSV
    expense_inflation_pct = 0.0
    if can_use_economic_scenario:
        with st.expander("Economic scenario (benefit indexation & expense inflation)", expanded=True):
            use_index = st.checkbox(
                "Use S&P 500 proxy CSV for benefit return indexation",
                help="If off, index is flat (zero equity returns); benefits stay level in nominal terms.",
                key="run_use_index",
            )
            index_csv = st.text_input(
                "Index scenario CSV (columns: month, sp500_level for months 0..N)",
                key="run_index_csv",
            )
            expense_inflation_pct = st.number_input(
                "Expense annual inflation (%, not tied to S&P)",
                min_value=0.0,
                max_value=25.0,
                help="Applied monthly as (1 + annual)^(1/12) to maintenance expenses only.",
                key="run_expense_inflation_pct",
            )

    mc_enable = False
    mc_n_sims = 100
    mc_seed = 42
    mc_drift_pct = 6.0
    mc_vol_pct = 15.0
    mc_s0 = 100.0
    if can_use_monte_carlo:
        with st.expander("Monte Carlo (stochastic index assumption)", expanded=True):
            mc_enable = st.checkbox(
                "Enable Monte Carlo on index returns",
                help="Simulates index paths and reprices for each path. Mortality, curve, and expense inflation remain deterministic.",
                key="run_mc_enable",
            )
            mc_n_sims = st.number_input("Number of simulations", min_value=100, max_value=20000, step=100, key="run_mc_n_sims")
            mc_seed = st.number_input("Random seed", min_value=0, max_value=2_147_483_647, step=1, key="run_mc_seed")
            mc_drift_pct = st.number_input("Annual drift (%)", min_value=-50.0, max_value=50.0, step=0.1, key="run_mc_drift_pct")
            mc_vol_pct = st.number_input("Annual volatility (%)", min_value=0.0, max_value=200.0, step=0.1, key="run_mc_vol_pct")
            mc_s0 = st.number_input("Initial index level (S0)", min_value=0.01, step=1.0, key="run_mc_s0")

    run = st.button("Run pricing", type="primary")

    if run:
        try:
            adapter = get_product_adapter(selected_product)
            yc = _build_yield_curve(
                y_mode,  # type: ignore[arg-type]
                flat_rate=flat_rate,
                zero_csv=zero_csv,
                par_csv=par_csv,
                coupon_freq=coupon_freq,
            )
            mort, needs_vy = _build_mortality(
                m_mode,  # type: ignore[arg-type]
                product_type=selected_product,
                sex="male" if sex == "male" else "female",
                qx_csv=qx_csv,
                rp_xlsx=rp_xlsx,
                rp_out_csv=rp_out,
                mp_xlsx=mp_xlsx,
                mp_out_csv=mp_out,
            )
            vy: int | None = int(valuation_year) if needs_vy else None
            vy_inputs = int(valuation_year)
            idx_path = str(_resolve_path(index_csv)) if use_index else None
            expense_annual_inflation = float(expense_inflation_pct) / 100.0

            if selected_product == ProductType.TERM_LIFE:
                contract = tp.TermLifeContract(
                    issue_age=int(issue_age),
                    sex="male" if sex == "male" else "female",
                    death_benefit=float(benefit_annual),
                    monthly_premium=float(monthly_premium),
                    term_years=20,
                    premium_mode="level_monthly",
                    benefit_timing="eoy_death",
                )
            else:
                contract = sp.SPIAContract(
                    issue_age=int(issue_age),
                    sex="male" if sex == "male" else "female",
                    benefit_annual=float(benefit_annual),
                )

            expenses_arg: sp.ExpenseAssumptions | None = None
            if expense_mode == "manual":
                expenses_arg = sp.ExpenseAssumptions(
                    policy_expense_dollars=pol,
                    premium_expense_rate=prem_pct / 100.0,
                    monthly_expense_dollars=monthly_ex,
                )
                expenses_used = expenses_arg
            else:
                try:
                    expenses_used = sp.ExpenseAssumptions.load_from_csv(str(_resolve_path(expenses_csv)))
                except (FileNotFoundError, ValueError, KeyError):
                    expenses_used = sp.ExpenseAssumptions(0.0, 0.0, 0.0)

            res = adapter.price(
                contract=contract,
                yield_curve=yc,
                mortality=mort,
                horizon_age=int(horizon_age),
                spread=float(spread),
                valuation_year=vy,
                expenses=expenses_arg,
                expenses_csv_path=str(_resolve_path(expenses_csv)),
                index_scenario_csv_path=idx_path,
                expense_annual_inflation=expense_annual_inflation,
            )
            _clear_dependent_state_on_pricing_change()
            st.session_state["pricing_res"] = res
            st.session_state["pricing_contract"] = contract
            st.session_state["pricing_product_type"] = selected_product.value
            st.session_state["pricing_run_id"] = int(st.session_state.get("pricing_run_id", 0)) + 1
            st.session_state["pricing_err"] = None
            st.session_state["pricing_meta"] = {
                "product_type": selected_product.value,
                "yield_mode": y_mode,
                "mortality_mode": m_mode,
                "expense_mode": expense_mode,
                "mc_enabled": bool(mc_enable and can_use_monte_carlo),
                "use_index": bool(use_index),
                "index_scenario_csv_path": idx_path,
            }
            st.session_state["pricing_run_inputs"] = {
                "sex": "male" if sex == "male" else "female",
                "issue_age": int(issue_age),
                "benefit_annual": float(benefit_annual),
                "horizon_age": int(horizon_age),
                "valuation_year": vy_inputs,
                "spread": float(spread),
                "expense_annual_inflation": float(expense_annual_inflation),
                "use_index": bool(use_index),
                "index_scenario_csv_path": idx_path,
                "mc_enabled": bool(mc_enable and can_use_monte_carlo),
                "mc_n_sims": int(mc_n_sims),
                "mc_seed": int(mc_seed),
                "mc_annual_drift": float(mc_drift_pct) / 100.0,
                "mc_annual_vol": float(mc_vol_pct) / 100.0,
                "mc_s0": float(mc_s0),
                "mc_base_settings_for_tail_risk": {
                    "annual_drift": float(mc_drift_pct) / 100.0,
                    "annual_vol": float(mc_vol_pct) / 100.0,
                    "seed": int(mc_seed),
                    "s0": float(mc_s0),
                },
                "term_length": term_choice,
                "term_premium_mode": premium_mode_choice,
                "term_benefit_timing": benefit_timing_choice,
                "term_monthly_premium": float(monthly_premium),
                "mortality_qx_csv": qx_csv,
                "mortality_rp_xlsx": rp_xlsx,
                "mortality_rp_out_csv": rp_out,
                "mortality_mp_xlsx": mp_xlsx,
                "mortality_mp_out_csv": mp_out,
            }
            st.session_state["pricing_excel_context"] = {
                "contract": contract,
                "yield_curve": yc,
                "mortality": mort,
                "horizon_age": int(horizon_age),
                "spread": float(spread),
                "valuation_year": vy_inputs,
                "expenses": expenses_used,
                "yield_mode": y_mode,
                "mortality_mode": m_mode,
                "expense_mode": expense_mode,
            }

            # --- Monte Carlo (run before Excel so MC sheet can be embedded) ---
            mc_snap_for_excel: MCExcelSnapshot | None = None
            if mc_enable and can_use_monte_carlo:
                mc = adapter.price_monte_carlo(
                    contract=contract,
                    yield_curve=yc,
                    mortality=mort,
                    horizon_age=int(horizon_age),
                    spread=float(spread),
                    valuation_year=vy,
                    expenses=expenses_arg,
                    expenses_csv_path=str(_resolve_path(expenses_csv)),
                    expense_annual_inflation=expense_annual_inflation,
                    n_sims=int(mc_n_sims),
                    annual_drift=float(mc_drift_pct) / 100.0,
                    annual_vol=float(mc_vol_pct) / 100.0,
                    seed=int(mc_seed),
                    s0=float(mc_s0),
                )
                st.session_state["pricing_mc"] = mc
                st.session_state["pricing_mc_params"] = {
                    "annual_drift": float(mc_drift_pct) / 100.0,
                    "annual_vol": float(mc_vol_pct) / 100.0,
                    "s0": float(mc_s0),
                    "n_sims": int(mc_n_sims),
                    "seed": int(mc_seed),
                }
                mc_snap_for_excel = mc_excel_snapshot_from_result(
                    mc,
                    annual_drift=float(mc_drift_pct) / 100.0,
                    annual_vol=float(mc_vol_pct) / 100.0,
                    s0=float(mc_s0),
                )
            else:
                st.session_state.pop("pricing_mc", None)
                st.session_state.pop("pricing_mc_params", None)

            # --- Excel workbook (built after MC so MC_Summary sheet can be included) ---
            _refresh_pricing_excel_workbook_in_session()
        except Exception as e:
            _clear_dependent_state_on_pricing_change()
            st.session_state["pricing_err"] = repr(e)
            st.session_state["pricing_res"] = None
            st.session_state.pop("pricing_product_type", None)
            st.session_state.pop("pricing_run_inputs", None)
            st.session_state.pop("pricing_excel_context", None)
            st.session_state.pop("pricing_xlsx_bytes", None)
            st.session_state.pop("pricing_xlsx_built_error", None)
            st.session_state.pop("pricing_mc", None)
            st.session_state.pop("pricing_mc_params", None)
            st.session_state.pop("pricing_xlsx_has_mc", None)
            st.session_state.pop("pricing_xlsx_has_alm", None)

    err = st.session_state.get("pricing_err")
    res = st.session_state.get("pricing_res")
    contract_state = st.session_state.get("pricing_contract")

    if err:
        st.error(err)
    if res is not None and contract_state is not None:
        st.success("Pricing completed.")
        meta = st.session_state.get("pricing_meta") or {}

        product_raw = str(meta.get("product_type", ProductType.SPIA.value))
        product_type = ProductType(product_raw) if product_raw in {p.value for p in ProductType} else ProductType.SPIA
        m1, m2, m3, m4 = st.columns(4)
        metrics = get_pricing_metrics(product_type, res)
        for col, metric in zip((m1, m2, m3, m4), metrics):
            formatted = f"${metric.value:,.0f}" if metric.is_money else f"{metric.value:,.0f}"
            col.metric(metric.label, formatted)

        st.caption(
            f"Yield: {meta.get('yield_mode')}; mortality: {meta.get('mortality_mode')}; "
            f"expenses: {meta.get('expense_mode')}."
        )
        mc_res = st.session_state.get("pricing_mc")
        if mc_res is not None:
            st.subheader("Monte Carlo summary (index-path uncertainty)")
            a1, a2, a3, a4 = st.columns(4)
            a1.metric("Mean premium", f"${mc_res.premium_mean:,.0f}")
            a2.metric("Median premium", f"${mc_res.premium_median:,.0f}")
            a3.metric("P5 premium", f"${mc_res.premium_p05:,.0f}")
            a4.metric("P95 premium", f"${mc_res.premium_p95:,.0f}")
            st.caption(f"Simulations: {mc_res.n_sims:,}")
            hist_counts, hist_edges = np.histogram(mc_res.single_premium, bins=40)
            hist_df = pd.DataFrame(
                {
                    "premium_bin_mid": 0.5 * (hist_edges[:-1] + hist_edges[1:]),
                    "count": hist_counts,
                }
            ).set_index("premium_bin_mid")
            st.line_chart(_round_for_visuals(hist_df))

        df = _result_dataframe(res, contract_state)
        st.subheader("Month-by-month projection")
        df_display = _round_for_visuals(df)
        st.dataframe(
            df_display,
            use_container_width=True,
            height=360,
            column_config=_number_cols_no_decimals(df_display),
        )
        csv_bytes = df_display.to_csv(index=False).encode("utf-8")
        c_dl1, c_dl2 = st.columns(2)
        with c_dl1:
            st.download_button(
                "Download projection CSV",
                data=csv_bytes,
                file_name=get_product_ui_config(product_type).projection_csv_filename,
                mime="text/csv",
            )
        with c_dl2:
            st.caption("Excel download moved to the Excel Replicator section.")
        ctx = st.session_state.get("pricing_excel_context") or {}
        expenses = ctx.get("expenses")
        _render_pricing_run_charts(res, contract_state, expenses)


def _read_workbook_cell_float(ws: Any, coord: str) -> float | None:
    v = ws[coord].value
    if v is None:
        return None
    try:
        x = float(v)
    except (TypeError, ValueError):
        return None
    if not np.isfinite(x):
        return None
    return x


def _alm_workbook_mirror_snapshot(
    alm: sp.ALMResult,
    asm: sp.ALMAssumptions | None,
    *,
    initial_asset_market_value: float | None,
) -> ALMExcelSnapshot | None:
    """Same truncation/downsample as ``build_workbook_from_spec`` (must match embedded ALM path)."""
    if asm is None:
        return None
    try:
        raw = alm_excel_snapshot_from_result(
            alm,
            asm,
            initial_asset_market_value=initial_asset_market_value,
        )
        ds = alm_excel_downsample_snapshot(raw, int(ALM_ENGINE_STEP_MONTHS))
        return alm_excel_truncate_snapshot(ds, int(ALM_EXCEL_PATH_MONTH_CAP))
    except Exception:
        return None


def _alm_modelcheck_key_assets_surplus_df(
    *,
    alm: sp.ALMResult,
    xlsx_bytes: bytes | None,
    dr: int = ALM_PROJECTION_FIRST_DATA_ROW,
    mirror_snap: ALMExcelSnapshot | None = None,
) -> pd.DataFrame:
    """
    ModelCheck-style table: Python path vs workbook (ALM_Projection).

    Uses the same truncated ALM snapshot as the Excel export when ``mirror_snap`` is provided so the
    Python column matches ModelCheck column B and the first ``ALM_EXCEL_PATH_MONTH_CAP`` rows on the sheet.

    Surplus from Excel is read as **C−D−E** when cached values exist (then **F** is not used). If the
    workbook has no cached results for those cells (typical before a full Excel recalc), the expected
    Excel column is **NaN** — do not treat that as a match. New downloads embed snapshot caches on
    **ALM_Projection** C–F and per-bucket columns **H+** so ``=SUM(bucket…)`` matches **C** under
    ``data_only`` until Excel recalculates (recalc may still refresh values from **ALM_Engine** formulas).
    """
    if mirror_snap is not None:
        a_mv = np.asarray(mirror_snap.asset_market_value, dtype=float)
        l_pv = np.asarray(mirror_snap.liability_pv, dtype=float)
        debt_b = np.asarray(mirror_snap.borrowing_balance, dtype=float)
        surp = a_mv - l_pv - debt_b
    else:
        a_mv = np.asarray(alm.asset_market_value, dtype=float)
        surp = np.asarray(alm.surplus, dtype=float)
    n = int(a_mv.size)
    if n < 1:
        return pd.DataFrame()

    n_rows = int(min(ALM_EXCEL_PATH_MONTH_CAP, n))
    if n_rows < 1:
        return pd.DataFrame()
    last_lab = f"ALM asset MV (month {n_rows} on sheet)"
    last_s_lab = f"ALM surplus (month {n_rows} on sheet)"
    if n_rows == 1:
        asset_specs: list[tuple[int, int, str]] = [(0, 0, "ALM asset MV (month 1 on sheet)")]
        surp_specs: list[tuple[int, int, str]] = [(0, 0, "ALM surplus (month 1 on sheet)")]
    else:
        asset_specs = [
            (0, 0, "ALM asset MV (month 1 on sheet)"),
            (n_rows - 1, n_rows - 1, last_lab),
        ]
        surp_specs = [
            (0, 0, "ALM surplus (month 1 on sheet)"),
            (n_rows - 1, n_rows - 1, last_s_lab),
        ]

    ws = None
    if isinstance(xlsx_bytes, bytes) and xlsx_bytes:
        try:
            wb = load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
            if ALM_SHEET_NAME in wb.sheetnames:
                ws = wb[ALM_SHEET_NAME]
        except Exception:
            ws = None

    rows: list[dict[str, Any]] = []
    for excel_off, py_idx, lab in asset_specs:
        py = float(a_mv[py_idx])
        r = dr + excel_off
        ex: float | None = None
        if ws is not None:
            ex = _read_workbook_cell_float(ws, f"C{r}")
        if ex is None:
            # Do not substitute Python: empty formula caches would fake a match (openpyxl data_only).
            ex = float("nan") if ws is not None else py
        rows.append(
            {
                "Metric": lab,
                "Python snapshot": py,
                "Expected Excel value (after recalc)": ex,
                "Difference (Excel - Python)": float(ex - py) if np.isfinite(ex) else float("nan"),
            }
        )

    for excel_off, py_idx, lab in surp_specs:
        py = float(surp[py_idx])
        r = dr + excel_off
        ex: float | None = None
        if ws is not None:
            c = _read_workbook_cell_float(ws, f"C{r}")
            d = _read_workbook_cell_float(ws, f"D{r}")
            e_b = _read_workbook_cell_float(ws, f"E{r}")
            if c is not None and d is not None and e_b is not None:
                ex = float(c - d - e_b)
            else:
                ex = _read_workbook_cell_float(ws, f"F{r}")
        if ex is None:
            ex = float("nan") if ws is not None else py
        rows.append(
            {
                "Metric": lab,
                "Python snapshot": py,
                "Expected Excel value (after recalc)": ex,
                "Difference (Excel - Python)": float(ex - py) if np.isfinite(ex) else float("nan"),
            }
        )

    return pd.DataFrame(rows)


def _render_excel_replicator() -> None:
    st.header("Excel Replicator")
    st.caption("Download the formula workbook and review parity metrics aligned with the workbook ModelCheck sheet.")

    res = st.session_state.get("pricing_res")
    contract_state = st.session_state.get("pricing_contract")
    if res is None or contract_state is None:
        st.info("Run pricing first in the Pricing Run section to populate the Excel Replicator.")
        return

    meta = st.session_state.get("pricing_meta") or {}
    product_raw = str(meta.get("product_type", ProductType.SPIA.value))
    product_type = ProductType(product_raw) if product_raw in {p.value for p in ProductType} else ProductType.SPIA

    _ensure_excel_workbook_includes_current_alm()

    m1, m2, m3, m4 = st.columns(4)
    metrics = get_pricing_metrics(product_type, res)
    for col, metric in zip((m1, m2, m3, m4), metrics):
        formatted = f"${metric.value:,.0f}" if metric.is_money else f"{metric.value:,.0f}"
        col.metric(f"Python {metric.label.lower()}", formatted)

    modelcheck = pd.DataFrame(
        [
            {
                "Metric": "PV benefits",
                "Python snapshot": float(res.pv_benefit),
                "Expected Excel value (after recalc)": float(res.pv_benefit),
                "Difference (Excel - Python)": 0.0,
            },
            {
                "Metric": "PV monthly expenses",
                "Python snapshot": float(res.pv_monthly_expenses),
                "Expected Excel value (after recalc)": float(res.pv_monthly_expenses),
                "Difference (Excel - Python)": 0.0,
            },
            {
                "Metric": "PV monthly total (ben+exp)",
                "Python snapshot": float(res.pv_benefit + res.pv_monthly_expenses),
                "Expected Excel value (after recalc)": float(res.pv_benefit + res.pv_monthly_expenses),
                "Difference (Excel - Python)": 0.0,
            },
            {
                "Metric": "Single premium",
                "Python snapshot": float(res.single_premium),
                "Expected Excel value (after recalc)": float(res.single_premium),
                "Difference (Excel - Python)": 0.0,
            },
            {
                "Metric": "Annuity factor",
                "Python snapshot": float(res.annuity_factor),
                "Expected Excel value (after recalc)": float(res.annuity_factor),
                "Difference (Excel - Python)": 0.0,
            },
        ]
    )

    st.subheader("ModelCheck parity dashboard")
    modelcheck_display = _round_for_visuals(modelcheck)
    st.dataframe(
        modelcheck_display,
        use_container_width=True,
        hide_index=True,
        column_config=_number_cols_no_decimals(modelcheck_display),
    )
    st.caption(
        f"Workbook references: PV benefits `{LIABILITY_SHEET_NAME}!X4`, PV monthly expenses `{LIABILITY_SHEET_NAME}!X5`, "
        f"PV monthly total `{LIABILITY_SHEET_NAME}!X7`, single premium `{LIABILITY_SHEET_NAME}!X8`, annuity factor `{LIABILITY_SHEET_NAME}!X6`."
    )
    st.caption(
        "After opening the workbook and recalculating, the ModelCheck tab differences should be near zero "
        "if Inputs match this run (especially spread B9 and valuation year)."
    )

    alm_chk = st.session_state.get("alm_last")
    alm_chk_rid = st.session_state.get("alm_last_pricing_run_id")
    pr_chk_rid = st.session_state.get("pricing_run_id")
    if isinstance(alm_chk, sp.ALMResult) and alm_chk_rid == pr_chk_rid:
        st.subheader("ModelCheck — ALM (assets & surplus)")
        xb_mc = st.session_state.get("pricing_xlsx_bytes")
        asm_chk = st.session_state.get("alm_last_assumptions")
        aum_chk = st.session_state.get("alm_last_initial_asset_market_value")
        aum_opt = float(aum_chk) if aum_chk is not None else None
        mirror = (
            _alm_workbook_mirror_snapshot(
                alm_chk,
                asm_chk if isinstance(asm_chk, sp.ALMAssumptions) else None,
                initial_asset_market_value=aum_opt,
            )
            if isinstance(asm_chk, sp.ALMAssumptions)
            else None
        )
        alm_mc_df = _alm_modelcheck_key_assets_surplus_df(
            alm=alm_chk,
            xlsx_bytes=xb_mc if isinstance(xb_mc, bytes) else None,
            mirror_snap=mirror,
        )
        if not alm_mc_df.empty:
            alm_mc_disp = _round_for_visuals(alm_mc_df)
            st.dataframe(
                alm_mc_disp,
                use_container_width=True,
                hide_index=True,
                column_config=_number_cols_no_decimals(alm_mc_disp),
            )
            n_mon = int(np.asarray(alm_chk.asset_market_value).size)
            n_on_sheet = int(min(ALM_EXCEL_PATH_MONTH_CAP, n_mon))
            lr = ALM_PROJECTION_FIRST_DATA_ROW + n_on_sheet - 1
            st.caption(
                f"Workbook **{ALM_SHEET_NAME}** / **{ALM_ENGINE_SHEET}** show the **first {n_on_sheet}** monthly ALM steps "
                f"(cap {ALM_EXCEL_PATH_MONTH_CAP}; Python may have more months). Rows **{ALM_PROJECTION_FIRST_DATA_ROW}**–**{lr}**. "
                f"**C** = SUM buckets; **D** from **{LIABILITY_SHEET_NAME}**; **F** = C−D−E. "
                "Parity uses cached **C−D−E** (embedded on export). After a full recalc in Excel, "
                "saved values may differ if formulas diverge from Python; re-download to reset caches."
            )

    # --- Monte Carlo distribution dashboard ---
    mc_res = st.session_state.get("pricing_mc")
    mc_params = st.session_state.get("pricing_mc_params") or {}
    if mc_res is not None:
        st.divider()
        st.subheader("Monte Carlo distribution statistics")
        n_sims_disp = mc_params.get("n_sims", mc_res.n_sims)
        drift_disp = mc_params.get("annual_drift", 0.0)
        vol_disp = mc_params.get("annual_vol", 0.0)
        s0_disp = mc_params.get("s0", 100.0)
        st.caption(
            f"{n_sims_disp:,} simulations | GBM drift {drift_disp * 100:.1f}% | "
            f"vol {vol_disp * 100:.1f}% | S\u2080 {s0_disp:.2f} | "
            "Mortality, yield curve, and expense inflation are deterministic across paths."
        )

        _mc_metrics: list[tuple[str, np.ndarray]] = [
            ("Single Premium ($)", mc_res.single_premium),
            ("PV Benefit ($)", mc_res.pv_benefit),
            ("PV Monthly Expenses ($)", mc_res.pv_monthly_expenses),
            ("PV Monthly Total ($)", mc_res.pv_monthly_total),
            ("Annuity Factor", mc_res.annuity_factor),
        ]
        stat_rows = []
        for name, arr in _mc_metrics:
            stat_rows.append(
                {
                    "Metric": name,
                    "Mean": float(np.mean(arr)),
                    "Std Dev": float(np.std(arr)),
                    "P5": float(np.percentile(arr, 5)),
                    "P25": float(np.percentile(arr, 25)),
                    "Median": float(np.median(arr)),
                    "P75": float(np.percentile(arr, 75)),
                    "P95": float(np.percentile(arr, 95)),
                }
            )
        stats_df = pd.DataFrame(stat_rows)
        stats_display = _round_for_visuals(stats_df)
        st.dataframe(
            stats_display,
            use_container_width=True,
            hide_index=True,
            column_config=_number_cols_no_decimals(stats_display),
        )

        st.markdown("**Premium & key metric distributions**")
        ch1, ch2 = st.columns(2)

        def _hist_df(arr: np.ndarray, n_bins: int = 35) -> pd.DataFrame:
            counts, edges = np.histogram(arr, bins=n_bins)
            mids = 0.5 * (edges[:-1] + edges[1:])
            return pd.DataFrame({"bin": np.rint(mids), "count": counts}).set_index("bin")

        with ch1:
            st.markdown("Single premium")
            st.bar_chart(_hist_df(mc_res.single_premium))
        with ch2:
            st.markdown("PV benefit")
            st.bar_chart(_hist_df(mc_res.pv_benefit))

        ch3, ch4 = st.columns(2)
        with ch3:
            st.markdown("Annuity factor")
            st.bar_chart(_hist_df(mc_res.annuity_factor))
        with ch4:
            st.markdown("PV monthly total")
            st.bar_chart(_hist_df(mc_res.pv_monthly_total))

        st.caption(
            "The MC_Summary sheet in the downloaded workbook contains the same statistics table "
            "and a premium distribution chart embedded as an Excel bar chart."
        )
    else:
        st.info(
            "Monte Carlo was not enabled for this run. Enable it in the Pricing Run section "
            "and re-run to see distribution statistics here and in the Excel workbook."
        )

    st.divider()
    xb = st.session_state.get("pricing_xlsx_bytes")
    xlsx_has_mc: bool = st.session_state.get("pricing_xlsx_has_mc", False)
    xlsx_has_alm: bool = st.session_state.get("pricing_xlsx_has_alm", False)
    if isinstance(xb, bytes) and xb:
        if xlsx_has_mc:
            st.success(
                "Workbook includes **MC_Summary** sheet with distribution statistics table and premium histogram chart.",
                icon="✅",
            )
        else:
            st.warning(
                "Workbook does **not** include MC_Summary — MC was disabled or not run when this workbook was built. "
                "Enable Monte Carlo in Pricing Run and click **Run pricing** again to regenerate.",
                icon="⚠️",
            )
        if xlsx_has_alm:
            st.success(
                f"Workbook includes **{LIABILITY_SHEET_NAME}**, **ALM_Projection** / **ALM_Engine** / **{ALM_ENGINE_FIELD_GUIDE_SHEET}** "
                f"(first {ALM_EXCEL_PATH_MONTH_CAP} months of the ALM path) and **Dashboard** ALM charts.",
                icon="✅",
            )
        else:
            st.info("This workbook file does not yet include **ALM_Projection** — run ALM, then return here to refresh the download.")
        suffix_parts: list[str] = []
        if xlsx_has_mc:
            suffix_parts.append("MC_Summary")
        if xlsx_has_alm:
            suffix_parts.append("ALM")
        mc_label = (" + " + " + ".join(suffix_parts)) if suffix_parts else ""
        help_bits = ["ModelCheck parity vs Python pricing"]
        if xlsx_has_mc:
            help_bits.append("MC statistics chart")
        if xlsx_has_alm:
            help_bits.append("ALM path sheet and dashboard charts")
        st.download_button(
            f"Download Excel recalculation workbook{mc_label}",
            data=xb,
            file_name=get_product_ui_config(product_type).recalc_workbook_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Workbook includes " + ", ".join(help_bits) + ".",
            type="primary",
        )
    elif st.session_state.get("pricing_xlsx_built_error"):
        st.error(f"Excel export unavailable: {st.session_state['pricing_xlsx_built_error']}")
    else:
        st.warning("Excel workbook not available yet for this run.")


def _render_alm_section() -> None:
    st.header("Asset–liability management (ALM)")
    st.caption(
        "Dynamic Treasury ladder + cash versus priced SPIA outflows. Earned rate on assets and liability discounting use "
        "the same zero curve + credit spread as **Pricing Run** for consistency. "
        "Rebalancing uses a drift band versus target weights on the review months you choose."
    )

    res = st.session_state.get("pricing_res")
    contract_state = st.session_state.get("pricing_contract")
    ctx = st.session_state.get("pricing_excel_context") or {}
    yc = ctx.get("yield_curve")
    spr = float(ctx.get("spread", 0.0))

    if (
        res is None
        or contract_state is None
        or not isinstance(yc, sp.YieldCurve)
    ):
        st.info("Run **Pricing Run** first. ALM anchors on that liability path, curve, and spread.")
        return

    # Allow the ALM engine to run against either the Pricing Run deterministic baseline,
    # or against a single Monte Carlo index path scenario (liability PV + discounting).
    scenario_source: str = "Base (Pricing Run deterministic)"
    mc_params = st.session_state.get("pricing_mc_params") or {}
    mc_n_sims = int(mc_params.get("n_sims", 0) or 0)
    mc_seed = int(mc_params.get("seed", 42) or 42)
    mc_scenario_idx: int = 0
    if mc_n_sims > 0:
        scenario_source = st.selectbox(
            "ALM pricing scenario (for liability PV and discounting)",
            options=["Base (Pricing Run deterministic)", "MC simulation (single path)"],
            index=0,
        )
        if scenario_source == "MC simulation (single path)":
            mc_scenario_idx = st.number_input(
                "MC simulation index (0-based)",
                min_value=0,
                max_value=max(0, mc_n_sims - 1),
                value=0,
                step=1,
            )
    else:
        st.caption("MC scenario selection is unavailable because Pricing Run MC inputs are missing.")

    base_spec = sp.alm_default_allocation_spec()
    n_bk = len(base_spec.buckets)
    # Initialize allocation widget state once; keyed widgets then read from session state only.
    for i in range(n_bk):
        k = f"alm_alloc_{i}"
        if k not in st.session_state:
            st.session_state[k] = float(round(base_spec.weights[i] * 100.0, 2))
    # Apply optimized weights safely on next rerun (avoid mutating active widget keys mid-run).
    pending_alloc = st.session_state.pop("alm_alloc_pending", None)
    if isinstance(pending_alloc, (list, tuple, np.ndarray)) and len(pending_alloc) == n_bk:
        try:
            for i, wi in enumerate(np.asarray(pending_alloc, dtype=float)):
                st.session_state[f"alm_alloc_{i}"] = float(wi * 100.0)
        except Exception:
            pass
    opt_notice = st.session_state.pop("alm_opt_notice", None)
    if isinstance(opt_notice, dict):
        msg = str(opt_notice.get("message", ""))
        level = str(opt_notice.get("level", "info"))
        if level == "success":
            st.success(msg)
        elif level == "warning":
            st.warning(msg)
        else:
            st.info(msg)

    with st.expander("Target allocation (% weights, must sum to 100%)", expanded=True):
        cols = st.columns(min(n_bk, 6))
        raw: list[float] = []
        for i, b in enumerate(base_spec.buckets):
            with cols[i % len(cols)]:
                raw.append(
                    float(
                        st.number_input(
                            f"{b.name} %",
                            min_value=0.0,
                            max_value=100.0,
                            step=0.5,
                            key=f"alm_alloc_{i}",
                        )
                    )
                )
        norm_run = st.checkbox("Normalize percentages to 100% on run", value=True)
        ws = np.array(raw, dtype=float) / 100.0
        s = float(np.sum(ws))
        if s <= 0.0:
            st.error("Allocation must include positive weights.")
        elif abs(s - 1.0) > 1e-3 and not norm_run:
            st.warning(f"Weights currently sum to {s * 100:.2f}%. Enable normalization or adjust inputs.")
        elif abs(s - 1.0) > 1e-3 and norm_run:
            st.info(f"Weights sum to {s * 100:.2f}% — will scale to 100% on run.")

    with st.expander("Rebalancing, flows, liquidity definition", expanded=True):
        c1, c2 = st.columns(2)
        with c1:
            band_pct = st.slider("Drift band vs targets (± share of AUM)", 0.5, 20.0, 5.0, 0.5)
            freq_m = st.number_input("Check rebalance every N months", min_value=1, value=1, step=1)
            near_liq_y = st.number_input(
                "Near-liquid residual maturity (years)",
                min_value=0.0,
                max_value=3.0,
                value=0.25,
                step=0.05,
                help="Liquidity buffer counts cash plus bond market value in buckets with residual maturity here or below.",
            )
            borrow_policy = st.selectbox(
                "Borrowing policy",
                options=["borrow_before_selling", "borrow_after_assets_insufficient"],
                index=1,
                format_func=lambda x: {
                    "borrow_before_selling": "Always borrow before selling assets",
                    "borrow_after_assets_insufficient": "Borrow only when asset portfolio is insufficient",
                }[x],
            )
            borrow_rate_mode = st.selectbox(
                "Borrowing rate basis",
                options=["scenario_linked", "fixed"],
                index=0,
                format_func=lambda x: {
                    "scenario_linked": "Scenario-linked (selected tenor rate + spread)",
                    "fixed": "Fixed annual borrowing rate",
                }[x],
            )
            is_scenario_linked = borrow_rate_mode == "scenario_linked"
            borrow_rate_tenor = st.selectbox(
                "Scenario-linked borrowing tenor",
                options=[0.25, 0.5, 1.0, 2.0, 3.0, 5.0],
                index=2,
                format_func=lambda x: f"{x:g}Y",
                help="Curve tenor used to derive borrowing base rate in scenario-linked mode.",
                disabled=not is_scenario_linked,
            )
            # Logical default: 1Y curve+spread plus 100 bps floor at 3%.
            df_t = float(yc.discount_factors(np.array([float(borrow_rate_tenor)], dtype=float), spread=spr)[0])
            base_t = -np.log(max(df_t, 1e-15)) / float(borrow_rate_tenor)
            borrow_rate_default_pct = float(max(0.03, base_t + 0.01) * 100.0)
            borrow_spread_bps = st.number_input(
                "Borrowing spread over selected tenor scenario rate (bps)",
                min_value=0.0,
                max_value=2000.0,
                value=100.0,
                step=5.0,
                help="Used when borrowing rate basis is scenario-linked.",
                disabled=not is_scenario_linked,
            )
            borrow_rate_pct = st.number_input(
                "Fixed borrowing rate (annual, %)",
                min_value=0.0,
                max_value=50.0,
                value=round(borrow_rate_default_pct, 2),
                step=0.1,
                help="Used only when borrowing rate basis is fixed.",
                disabled=is_scenario_linked,
            )
        with c2:
            rebalance_policy = st.selectbox(
                "Rebalance policy",
                options=["liquidity_only", "full_target"],
                index=0,
                format_func=lambda x: {
                    "liquidity_only": "Hold to maturity bias (sell only for liquidity shortfall)",
                    "full_target": "Full target rebalance (trade back to weights when drift breaches band)",
                }[x],
            )
            reinvest = st.selectbox(
                "Reinvest matured principal",
                options=["hold_cash", "pro_rata"],
                index=1,
                format_func=lambda x: {
                    "hold_cash": "Keep in cash until band rebalance (or next review)",
                    "pro_rata": "Re-deploy into bonds pro-rata to bond targets (excess over cash target)",
                }[x],
            )
            disinvest = st.selectbox(
                "Disinvest if cash is short after outflows",
                options=["shortest_first", "pro_rata"],
                index=0,
                format_func=lambda x: {
                    "shortest_first": "Shortest residual maturity first",
                    "pro_rata": "Pro-rata across bond buckets",
                }[x],
            )

    aum0 = st.number_input(
        "Initial asset market value ($)",
        min_value=0.0,
        value=float(res.single_premium),
        step=10_000.0,
        help="Usually the priced single premium invested at issue.",
    )

    # Persist the current ALM selection so What-if and diagnostics can reflect the user's latest inputs
    # even before they click "Run ALM projection".
    try:
        ws_run_current = ws.copy()
        if norm_run and float(np.sum(ws_run_current)) > 0.0:
            ws_run_current = ws_run_current / float(np.sum(ws_run_current))
        alloc_current = sp.ALMAllocationSpec(buckets=base_spec.buckets, weights=ws_run_current)
        asm_current = sp.ALMAssumptions(
            allocation=alloc_current,
            rebalance_band=float(band_pct) / 100.0,
            rebalance_frequency_months=int(freq_m),
            reinvest_rule=reinvest,  # type: ignore[arg-type]
            disinvest_rule=disinvest,  # type: ignore[arg-type]
            rebalance_policy=rebalance_policy,  # type: ignore[arg-type]
            borrowing_policy=borrow_policy,  # type: ignore[arg-type]
            borrowing_rate_mode=borrow_rate_mode,  # type: ignore[arg-type]
            borrowing_rate_tenor_years=float(borrow_rate_tenor),
            borrowing_spread_annual=float(borrow_spread_bps) / 10000.0,
            borrowing_rate_annual=float(borrow_rate_pct) / 100.0,
            liquidity_near_liquid_years=float(near_liq_y),
        )
        st.session_state["alm_current_assumptions"] = asm_current
        st.session_state["alm_current_initial_asset_market_value"] = float(aum0)
    except Exception:
        # Don't block UI if current weights/band inputs are temporarily invalid while user is editing.
        st.session_state.pop("alm_current_assumptions", None)
        st.session_state.pop("alm_current_initial_asset_market_value", None)

    run_alm = st.button("Run ALM projection", type="primary")
    opt_col1, opt_col2 = st.columns([2, 1])
    with opt_col1:
        opt_surplus_constraint = st.selectbox(
            "Optimization surplus constraint",
            options=["Path never negative", "Ending surplus non-negative"],
            index=1,
            help="Select whether optimization enforces surplus >= 0 at every month, or only at the ending month.",
        )
    with opt_col2:
        opt_samples = st.number_input(
            "Optimization samples",
            min_value=100,
            max_value=5000,
            value=1200,
            step=100,
            help="Requested random candidates; runtime controls may cap evaluated candidates for responsiveness.",
        )
    opt_objective = st.selectbox(
        "Optimization objective",
        options=["Balanced mix (diversified weights)", "Match liability KRD by tenor (fast screen + ALM)"],
        index=0,
        help=(
            "Balanced mix scores weights for diversification vs targets, then runs ALM on each candidate until caps. "
            "KRD match uses a larger random pool plus simplex coordinate refinement on the analytical hedge score, "
            "then runs ALM on more top candidates (higher time budget) for surplus feasibility."
        ),
    )
    run_alm_opt = st.button("Optimize allocation and run ALM")

    def _build_pricing_for_selected_scenario() -> sp.SPIAProjectionResult:
        if scenario_source != "MC simulation (single path)":
            return res
        mort = ctx.get("mortality")
        expenses = ctx.get("expenses")
        valuation_year = ctx.get("valuation_year")
        horizon_age = ctx.get("horizon_age")
        if not isinstance(mort, (sp.MortalityTableQx, sp.MortalityTableRP2014MP2016)):
            raise ValueError("Pricing Run mortality missing from session state.")
        if not isinstance(expenses, sp.ExpenseAssumptions):
            raise ValueError("Pricing Run expenses missing from session state.")

        n_months = int(res.months.size)
        idx_paths = sp.simulate_index_levels_gbm(
            n_sims=mc_n_sims,
            n_months=n_months,
            s0=float(mc_params.get("s0", 100.0) or 100.0),
            annual_drift=float(mc_params.get("annual_drift", 0.06) or 0.06),
            annual_vol=float(mc_params.get("annual_vol", 0.15) or 0.15),
            seed=mc_seed,
        )
        idx_one = idx_paths[int(mc_scenario_idx)]
        idx_levels_payment = np.asarray(idx_one[1:], dtype=float)
        idx_s0 = float(idx_one[0])
        return sp.price_spia_single_premium(
            contract=contract_state,
            yield_curve=yc,
            mortality=mort,
            horizon_age=int(horizon_age),
            spread=spr,
            valuation_year=int(valuation_year) if valuation_year is not None else None,
            expenses=expenses,
            expense_annual_inflation=float(res.expense_annual_inflation),
            index_s0=idx_s0,
            index_levels_payment=idx_levels_payment,
        )

    if run_alm:
        if aum0 <= 0.0:
            st.error("Initial assets must be positive.")
        elif float(np.sum(ws)) <= 0.0:
            st.error("Invalid allocation.")
        else:
            ws_run = ws.copy()
            if norm_run and float(np.sum(ws_run)) > 0:
                ws_run = ws_run / float(np.sum(ws_run))
            try:
                alloc = sp.ALMAllocationSpec(buckets=base_spec.buckets, weights=ws_run)
                asm = sp.ALMAssumptions(
                    allocation=alloc,
                    rebalance_band=float(band_pct) / 100.0,
                    rebalance_frequency_months=int(freq_m),
                    reinvest_rule=reinvest,  # type: ignore[arg-type]
                    disinvest_rule=disinvest,  # type: ignore[arg-type]
                    rebalance_policy=rebalance_policy,  # type: ignore[arg-type]
                    borrowing_policy=borrow_policy,  # type: ignore[arg-type]
                    borrowing_rate_mode=borrow_rate_mode,  # type: ignore[arg-type]
                    borrowing_rate_tenor_years=float(borrow_rate_tenor),
                    borrowing_spread_annual=float(borrow_spread_bps) / 10000.0,
                    borrowing_rate_annual=float(borrow_rate_pct) / 100.0,
                    liquidity_near_liquid_years=float(near_liq_y),
                )
                pricing_for_alm = _build_pricing_for_selected_scenario()
                out = sp.run_alm_projection(
                    pricing=pricing_for_alm,
                    yield_curve=yc,
                    spread=spr,
                    assumptions=asm,
                    initial_asset_market_value=float(aum0),
                )
                st.session_state["alm_last"] = out
                st.session_state["alm_last_assumptions"] = asm
                st.session_state["alm_last_initial_asset_market_value"] = float(aum0)
                st.session_state["alm_last_pricing_run_id"] = st.session_state.get("pricing_run_id")
                _invalidate_diagnostics_export()
                _refresh_pricing_excel_workbook_in_session()
                st.success("ALM projection complete.")
            except Exception as ex:
                st.error(f"ALM run failed: {ex!r}")

    if run_alm_opt:
        if aum0 <= 0.0:
            st.error("Initial assets must be positive.")
        else:
            try:
                pricing_for_alm = _build_pricing_for_selected_scenario()
                rng = np.random.default_rng(42)
                n_assets = len(base_spec.buckets)
                current_w = ws.copy()
                if float(np.sum(current_w)) > 0:
                    current_w = current_w / float(np.sum(current_w))
                candidates = [current_w, np.asarray(base_spec.weights, dtype=float)]
                # Add a conservative anchor.
                w_cons = np.zeros(n_assets, dtype=float)
                w_cons[0] = 0.20
                rem = 0.80 / max(1, n_assets - 1)
                w_cons[1:] = rem
                candidates.append(w_cons)
                # Structured tenor candidates around diversified bond ladders.
                tenors = np.array([float(b.tenor_years) for b in base_spec.buckets], dtype=float)
                bond_ten = np.clip(tenors[1:], 1e-9, None)
                for cash_w in [0.0, 0.05, 0.10, 0.20]:
                    for tilt in [-1.0, -0.5, 0.0, 0.5, 1.0]:
                        wb = bond_ten ** tilt
                        wb = wb / float(np.sum(wb))
                        w_try = np.concatenate(([cash_w], (1.0 - cash_w) * wb))
                        candidates.append(w_try)
                # Explicitly include near-even spreads to accelerate convergence to diversified mixes.
                for cash_w in [0.0, 0.05, 0.10]:
                    wb_even = np.full(n_assets - 1, (1.0 - cash_w) / float(max(1, n_assets - 1)), dtype=float)
                    candidates.append(np.concatenate(([cash_w], wb_even)))

                # Runtime guardrails to avoid very long runs on large horizons.
                if opt_surplus_constraint == "Ending surplus non-negative":
                    # Default mode is easier to satisfy; use tighter limits for better responsiveness.
                    max_eval = min(220, max(50, int(opt_samples)))
                    time_budget_sec = 6.0
                else:
                    max_eval = min(400, max(80, int(opt_samples)))
                    time_budget_sec = 12.0
                opt_krd_match = str(opt_objective).startswith("Match")
                if opt_krd_match:
                    # Allow more full ALM checks and wall-clock time for KRD matching.
                    time_budget_sec = min(32.0, float(time_budget_sec) + 14.0)
                    max_eval = min(520, int(max_eval) + 140)

                max_random = max(0, max_eval - len(candidates))
                # Random simplex samples (bounded by max_eval).
                alpha = np.ones(n_assets, dtype=float)
                for _ in range(min(int(opt_samples), max_random)):
                    candidates.append(rng.dirichlet(alpha))
                if opt_krd_match:
                    extra_draws = min(950, max(320, int(opt_samples) * 2 + 150))
                    for _ in range(extra_draws):
                        candidates.append(rng.dirichlet(alpha))

                key_tenors_opt = np.array(
                    [float(b.tenor_years) for b in base_spec.buckets[1:] if float(b.tenor_years) > 1e-12],
                    dtype=float,
                )
                bond_tenors_opt = key_tenors_opt.copy()

                tenor_axis = np.array([float(b.tenor_years) for b in base_spec.buckets], dtype=float)
                target_tenor = float(np.median(tenor_axis[1:])) if tenor_axis.size > 1 else 0.0
                best_score = float("inf")
                best_end_surplus = -float("inf")
                best_w: np.ndarray | None = None
                best_out: sp.ALMResult | None = None
                best_min_surplus = -float("inf")
                best_fallback_w: np.ndarray | None = None
                best_fallback_out: sp.ALMResult | None = None
                best_krd_mismatch = float("nan")

                start_t = time.perf_counter()
                eval_count = 0

                def _run_one_alm_candidate(w_try: np.ndarray, *, objective_score: float | None) -> None:
                    nonlocal best_score, best_end_surplus, best_w, best_out, eval_count
                    nonlocal best_min_surplus, best_fallback_w, best_fallback_out, best_krd_mismatch
                    w_norm = np.asarray(w_try, dtype=float)
                    s_wm = float(np.sum(w_norm))
                    if s_wm <= 1e-15:
                        return
                    w_norm = w_norm / s_wm
                    alloc_try = sp.ALMAllocationSpec(buckets=base_spec.buckets, weights=w_norm)
                    asm_try = sp.ALMAssumptions(
                        allocation=alloc_try,
                        rebalance_band=float(band_pct) / 100.0,
                        rebalance_frequency_months=int(freq_m),
                        reinvest_rule=reinvest,  # type: ignore[arg-type]
                        disinvest_rule=disinvest,  # type: ignore[arg-type]
                        rebalance_policy=rebalance_policy,  # type: ignore[arg-type]
                        borrowing_policy=borrow_policy,  # type: ignore[arg-type]
                        borrowing_rate_mode=borrow_rate_mode,  # type: ignore[arg-type]
                        borrowing_rate_tenor_years=float(borrow_rate_tenor),
                        borrowing_spread_annual=float(borrow_spread_bps) / 10000.0,
                        borrowing_rate_annual=float(borrow_rate_pct) / 100.0,
                        liquidity_near_liquid_years=float(near_liq_y),
                    )
                    out_try = sp.run_alm_projection(
                        pricing=pricing_for_alm,
                        yield_curve=yc,
                        spread=spr,
                        assumptions=asm_try,
                        initial_asset_market_value=float(aum0),
                    )
                    eval_count += 1

                    min_surp = float(np.min(np.asarray(out_try.surplus, dtype=float)))
                    if min_surp > best_min_surplus:
                        best_min_surplus = min_surp
                        best_fallback_w = w_norm.copy()
                        best_fallback_out = out_try

                    if opt_surplus_constraint == "Path never negative":
                        feasible = bool(np.all(np.asarray(out_try.surplus, dtype=float) >= -1e-6))
                    else:
                        feasible = bool(float(out_try.surplus[-1]) >= -1e-6)
                    if not feasible:
                        return

                    end_surplus = float(out_try.surplus[-1])
                    if opt_krd_match:
                        sc = float(objective_score) if objective_score is not None else float("inf")
                        if (
                            sc < best_score - 1e-15
                            or (abs(sc - best_score) <= 1e-15 and end_surplus > best_end_surplus)
                        ):
                            best_score = sc
                            best_end_surplus = end_surplus
                            best_w = w_norm.copy()
                            best_out = out_try
                            best_krd_mismatch = sc
                        return

                    w_eval = w_norm
                    w_bond = w_eval[1:] if w_eval.size > 1 else np.asarray([], dtype=float)
                    if w_bond.size > 0:
                        bond_sum = float(np.sum(w_bond))
                        if bond_sum > 1e-12:
                            w_bond_norm = w_bond / bond_sum
                            even_penalty = float(np.std(w_bond_norm))
                        else:
                            even_penalty = 1.0
                    else:
                        even_penalty = 0.0
                    tenor_score = float(np.dot(w_eval, tenor_axis))
                    tenor_dev_penalty = abs(tenor_score - target_tenor) / max(1.0, target_tenor)
                    long_penalty = float(w_eval[-1]) if w_eval.size > 1 else 0.0
                    concentration_penalty = float(np.max(w_eval))
                    score = (
                        1.00 * even_penalty
                        + 0.40 * tenor_dev_penalty
                        + 0.35 * long_penalty
                        + 0.25 * concentration_penalty
                    )
                    if (
                        score < best_score - 1e-12
                        or (abs(score - best_score) <= 1e-12 and end_surplus > best_end_surplus)
                    ):
                        best_score = score
                        best_end_surplus = end_surplus
                        best_w = w_norm.copy()
                        best_out = out_try

                if opt_krd_match:
                    if key_tenors_opt.size == 0:
                        st.error("KRD matching requires bond buckets with positive tenor.")
                    else:
                        liab_krd_vec = sp.liability_key_rate_durations_years(
                            yc,
                            float(spr),
                            np.asarray(pricing_for_alm.expected_total_cashflows, dtype=float),
                            np.asarray(pricing_for_alm.times_years, dtype=float),
                            key_tenors_opt,
                        )

                        def _analytical_krd_mismatch(wv: np.ndarray) -> float:
                            wv = np.maximum(np.asarray(wv, dtype=float), 0.0)
                            s = float(np.sum(wv))
                            if s <= 1e-15:
                                return float("inf")
                            wv = wv / s
                            ak = sp.initial_ladder_asset_key_rate_durations_years(
                                yc,
                                float(spr),
                                float(aum0),
                                wv,
                                bond_tenors_opt,
                                key_tenors_opt,
                            )
                            return float(sp.key_rate_duration_hedge_mismatch_score(ak, liab_krd_vec))

                        stage1_n = min(len(candidates), max(520, int(opt_samples) * 2 + 120))
                        krd_scored: list[tuple[float, np.ndarray]] = []
                        for w_try in candidates[:stage1_n]:
                            w_arr = np.asarray(w_try, dtype=float)
                            if float(np.sum(w_arr)) <= 1e-15:
                                continue
                            w_arr = w_arr / float(np.sum(w_arr))
                            sc = _analytical_krd_mismatch(w_arr)
                            krd_scored.append((sc, w_arr))
                        krd_scored.sort(key=lambda t: t[0])

                        refined_seen: set[tuple[float, ...]] = set()
                        seeds_to_refine: list[np.ndarray] = []
                        for _sc, wv in krd_scored[: min(18, len(krd_scored))]:
                            key = tuple(np.round(wv, 5).tolist())
                            if key in refined_seen:
                                continue
                            refined_seen.add(key)
                            seeds_to_refine.append(wv)
                            if len(seeds_to_refine) >= 10:
                                break
                        for w_seed in seeds_to_refine:
                            w_ref, sc_ref = sp.refine_weights_on_probability_simplex(
                                w_seed,
                                _analytical_krd_mismatch,
                                max_rounds=32,
                                transfer_fracs=(0.08, 0.05, 0.03, 0.02, 0.01, 0.006),
                            )
                            krd_scored.append((sc_ref, w_ref))

                        for rank in range(min(4, len(krd_scored))):
                            w_anchor = krd_scored[rank][1]
                            conc = 35.0 + 12.0 * float(rank)
                            for _ in range(72):
                                alpha_loc = np.maximum(np.asarray(w_anchor, dtype=float), 1e-4) * conc + 0.06
                                w_loc = rng.dirichlet(alpha_loc)
                                krd_scored.append((_analytical_krd_mismatch(w_loc), w_loc))

                        krd_scored.sort(key=lambda t: t[0])
                        top_m = min(58, max(38, max_eval // 2 + 8), len(krd_scored))
                        alm_picked: list[tuple[float, np.ndarray]] = []
                        seen_alm_weights: set[tuple[float, ...]] = set()
                        for sc_i, w_i in krd_scored:
                            keyw = tuple(np.round(np.asarray(w_i, dtype=float), 5).tolist())
                            if keyw in seen_alm_weights:
                                continue
                            seen_alm_weights.add(keyw)
                            alm_picked.append((sc_i, w_i))
                            if len(alm_picked) >= top_m:
                                break
                        for sc_i, w_i in alm_picked:
                            if (time.perf_counter() - start_t) >= time_budget_sec:
                                break
                            _run_one_alm_candidate(w_i, objective_score=sc_i)
                else:
                    for w_try in candidates:
                        if eval_count >= max_eval:
                            break
                        if (time.perf_counter() - start_t) >= time_budget_sec:
                            break
                        _run_one_alm_candidate(np.asarray(w_try, dtype=float), objective_score=None)

                if best_w is None or best_out is None:
                    if best_fallback_w is None or best_fallback_out is None:
                        st.warning(
                            "No feasible allocation found under the selected surplus constraint. "
                            "This can happen when constraints are too strict for current assumptions "
                            "(cashflows, borrowing policy/rate, rebalance policy, and curve)."
                        )
                    else:
                        st.session_state["alm_alloc_pending"] = np.asarray(best_fallback_w, dtype=float).tolist()
                        asm_best = sp.ALMAssumptions(
                            allocation=sp.ALMAllocationSpec(buckets=base_spec.buckets, weights=best_fallback_w),
                            rebalance_band=float(band_pct) / 100.0,
                            rebalance_frequency_months=int(freq_m),
                            reinvest_rule=reinvest,  # type: ignore[arg-type]
                            disinvest_rule=disinvest,  # type: ignore[arg-type]
                            rebalance_policy=rebalance_policy,  # type: ignore[arg-type]
                            borrowing_policy=borrow_policy,  # type: ignore[arg-type]
                            borrowing_rate_mode=borrow_rate_mode,  # type: ignore[arg-type]
                            borrowing_rate_tenor_years=float(borrow_rate_tenor),
                            borrowing_spread_annual=float(borrow_spread_bps) / 10000.0,
                            borrowing_rate_annual=float(borrow_rate_pct) / 100.0,
                            liquidity_near_liquid_years=float(near_liq_y),
                        )
                        st.session_state["alm_last"] = best_fallback_out
                        st.session_state["alm_last_assumptions"] = asm_best
                        st.session_state["alm_last_initial_asset_market_value"] = float(aum0)
                        st.session_state["alm_last_pricing_run_id"] = st.session_state.get("pricing_run_id")
                        _invalidate_diagnostics_export()
                        _refresh_pricing_excel_workbook_in_session()
                        st.session_state["alm_opt_notice"] = {
                            "level": "warning",
                            "message": (
                                "No feasible allocation found within runtime limits; showing nearest candidate "
                                "(highest minimum surplus). Target allocation inputs updated."
                            ),
                        }
                        st.rerun()
                else:
                    st.session_state["alm_alloc_pending"] = np.asarray(best_w, dtype=float).tolist()
                    asm_best = sp.ALMAssumptions(
                        allocation=sp.ALMAllocationSpec(buckets=base_spec.buckets, weights=best_w),
                        rebalance_band=float(band_pct) / 100.0,
                        rebalance_frequency_months=int(freq_m),
                        reinvest_rule=reinvest,  # type: ignore[arg-type]
                        disinvest_rule=disinvest,  # type: ignore[arg-type]
                        rebalance_policy=rebalance_policy,  # type: ignore[arg-type]
                        borrowing_policy=borrow_policy,  # type: ignore[arg-type]
                        borrowing_rate_mode=borrow_rate_mode,  # type: ignore[arg-type]
                        borrowing_rate_tenor_years=float(borrow_rate_tenor),
                        borrowing_spread_annual=float(borrow_spread_bps) / 10000.0,
                        borrowing_rate_annual=float(borrow_rate_pct) / 100.0,
                        liquidity_near_liquid_years=float(near_liq_y),
                    )
                    st.session_state["alm_last"] = best_out
                    st.session_state["alm_last_assumptions"] = asm_best
                    st.session_state["alm_last_initial_asset_market_value"] = float(aum0)
                    st.session_state["alm_last_pricing_run_id"] = st.session_state.get("pricing_run_id")
                    _invalidate_diagnostics_export()
                    _refresh_pricing_excel_workbook_in_session()
                    if opt_krd_match:
                        krd_msg = (
                            "Optimized allocation found (KRD screen: match asset key-rate sensitivities to liability "
                            f"by tenor; mean sq. rel. error {best_krd_mismatch:.4f}) and ALM projection completed. "
                            f"Weighted tenor: {float(np.dot(np.asarray(best_w, dtype=float), tenor_axis)):.2f}Y; "
                            f"ending surplus: ${float(best_out.surplus[-1]):,.0f}. "
                            "Target allocation inputs updated."
                        )
                    else:
                        krd_msg = (
                            "Optimized allocation found (balanced tenor spread with anti-concentration bias) "
                            "and ALM projection completed. "
                            f"Weighted tenor: {float(np.dot(np.asarray(best_w, dtype=float), tenor_axis)):.2f}Y; "
                            f"ending surplus: ${float(best_out.surplus[-1]):,.0f}. "
                            "Target allocation inputs updated."
                        )
                    st.session_state["alm_opt_notice"] = {"level": "success", "message": krd_msg}
                    st.rerun()
                st.caption(
                    f"Optimization evaluated {eval_count} ALM projection(s) "
                    f"(cap {max_eval}, time budget {time_budget_sec:.0f}s)."
                    + (" KRD mode ranks weights analytically first, then ALM-checks only the best few." if opt_krd_match else "")
                )
            except Exception as ex:
                st.error(f"ALM optimization failed: {ex!r}")

    last = st.session_state.get("alm_last")
    if isinstance(last, sp.ALMResult):
        st.subheader("ALM metrics (first month-end)")
        st.caption(
            "Path metrics are recorded after each month’s flows and trades. Scalar PV01 and durations are issue-time (initial portfolio)."
        )
        m1, m2, m3, m4, m5 = st.columns(5)
        with m1:
            fr0 = float(last.funding_ratio[0]) if last.funding_ratio.size else float("nan")
            st.metric("Funding ratio (month 1)", f"{fr0:.3f}")
        with m2:
            st.metric("Surplus ($)", f"${float(last.surplus[0]):,.0f}")
        with m3:
            st.metric("PV01 net ($/bp)", f"{float(last.pv01_net):,.0f}")
        with m4:
            st.metric("Duration gap (y)", f"{float(last.duration_gap):.2f}")
        with m5:
            lb0 = float(last.liquidity_buffer_months[0]) if last.liquidity_buffer_months.size else float("nan")
            st.metric("Liquidity buffer (mo)", f"{lb0:.2f}")

        st.subheader("Paths (attained age)")
        age_ax = contract_state.issue_age + last.times_years
        st.markdown("##### Asset market value and liability present value")
        st.line_chart(
            _round_for_visuals(
                pd.DataFrame(
                    {
                        "Asset market value": last.asset_market_value,
                        "Liability PV": last.liability_pv,
                    },
                    index=age_ax,
                )
            )
        )
        _alm_surplus_chart(age_ax, last.surplus)
        st.markdown("##### Liquidity buffer (months of mean monthly outflows)")
        st.line_chart(
            pd.DataFrame({"Liquidity buffer (months)": last.liquidity_buffer_months}, index=age_ax)
        )
        st.markdown("##### Borrowing balance")
        st.line_chart(
            _round_for_visuals(
                pd.DataFrame({"Borrowing balance": last.borrowing_balance}, index=age_ax)
            )
        )

        asm_vis = st.session_state.get("alm_last_assumptions")
        if isinstance(asm_vis, sp.ALMAssumptions):
            bucket_specs = list(asm_vis.allocation.buckets)
        else:
            bucket_specs = list(base_spec.buckets)
        # Keep all ALM legends/series in logical tenor order: cash, then shortest to longest tenor.
        order_idx = sorted(range(len(bucket_specs)), key=lambda i: float(bucket_specs[i].tenor_years))
        ordered_specs = [bucket_specs[i] for i in order_idx]
        ordered_names = [b.name for b in ordered_specs]

        bucket_df_raw = pd.DataFrame(last.bucket_asset_mv.T, columns=[b.name for b in bucket_specs], index=age_ax)
        bucket_df = bucket_df_raw.reindex(columns=ordered_names)
        st.markdown("**Bucket market values**")
        bucket_mv_long = (
            bucket_df.reset_index()
            .rename(columns={"index": "Attained age"})
            .melt(id_vars=["Attained age"], var_name="Asset type", value_name="Bucket market value")
        )
        bucket_mv_long["Asset type"] = pd.Categorical(bucket_mv_long["Asset type"], categories=ordered_names, ordered=True)
        bucket_mv_chart = (
            alt.Chart(bucket_mv_long)
            .mark_line()
            .encode(
                x=alt.X("Attained age:Q", title="Attained age"),
                y=alt.Y("Bucket market value:Q", title="Market value ($)"),
                color=alt.Color(
                    "Asset type:N",
                    title="Asset type",
                    sort=ordered_names,
                    legend=alt.Legend(orient="top", direction="horizontal", columns=len(ordered_names)),
                ),
                order=alt.Order("Asset type:N", sort="ascending"),
                tooltip=[
                    alt.Tooltip("Attained age:Q", format=".2f"),
                    alt.Tooltip("Asset type:N"),
                    alt.Tooltip("Bucket market value:Q", format=",.0f"),
                ],
            )
            .properties(height=320)
        )
        st.altair_chart(bucket_mv_chart.interactive(), use_container_width=True)

        st.markdown("##### Portfolio composition by asset type (%)")
        aum_series = pd.Series(last.asset_market_value, index=age_ax, dtype=float).replace(0.0, np.nan)
        weight_pct_df = bucket_df.div(aum_series, axis=0).fillna(0.0) * 100.0
        comp_df = (
            weight_pct_df.reset_index()
            .rename(columns={"index": "Attained age"})
            .melt(id_vars=["Attained age"], var_name="Asset type", value_name="Portfolio share (%)")
        )
        comp_df["Asset type"] = pd.Categorical(comp_df["Asset type"], categories=ordered_names, ordered=True)
        comp_chart = (
            alt.Chart(comp_df)
            .mark_area()
            .encode(
                x=alt.X("Attained age:Q", title="Attained age"),
                y=alt.Y(
                    "Portfolio share (%):Q",
                    stack=True,
                    title="Portfolio share (%)",
                    scale=alt.Scale(domain=[0, 100]),
                ),
                color=alt.Color(
                    "Asset type:N",
                    title="Asset type",
                    sort=ordered_names,
                    legend=alt.Legend(orient="top", direction="horizontal", columns=len(ordered_names)),
                ),
                order=alt.Order("Asset type:N", sort="ascending"),
            )
            .properties(height=320)
        )
        st.altair_chart(comp_chart.interactive(), use_container_width=True)

        # Yield decomposition: portfolio weighted yield plus per-asset-class contributions.
        # Contributions are weight * bucket annualized zero yield (incl. spread), shown in percentage points.
        tenors = np.array([float(b.tenor_years) for b in ordered_specs], dtype=float)
        bucket_yield = np.zeros_like(tenors, dtype=float)
        for i, T in enumerate(tenors):
            if T <= 1e-12:
                bucket_yield[i] = 0.0
            else:
                dff = float(yc.discount_factors(np.array([T], dtype=float), spread=spr)[0])
                bucket_yield[i] = -np.log(max(dff, 1e-15)) / T

        weight_df = bucket_df.div(aum_series, axis=0).fillna(0.0)
        contrib_pp_df = weight_df.mul(bucket_yield, axis=1) * 100.0
        total_yield_pct = contrib_pp_df.sum(axis=1).rename("Total portfolio yield (%)")
        total_line_df = total_yield_pct.reset_index().rename(columns={"index": "Attained age"})
        st.markdown("##### Total portfolio yield")
        yld_total = (
            alt.Chart(total_line_df)
            .mark_line(color="#1f77b4", strokeWidth=3)
            .encode(
                x=alt.X("Attained age:Q", title="Attained age"),
                y=alt.Y("Total portfolio yield (%):Q", title="Total yield (%)"),
                tooltip=[
                    alt.Tooltip("Attained age:Q", format=".2f"),
                    alt.Tooltip("Total portfolio yield (%):Q", format=".4f"),
                ],
            )
            .properties(height=320)
        )
        st.altair_chart(yld_total.interactive(), use_container_width=True)

        kpi_tbl = pd.DataFrame(
            {
                "PV01 assets": [last.pv01_assets],
                "PV01 liabilities": [last.pv01_liabilities],
                "Mac duration assets": [last.duration_assets_mac],
                "Mac duration liabilities": [last.duration_liabilities_mac],
            }
        )
        kpi_tbl_show = kpi_tbl.copy()
        kpi_tbl_show["PV01 assets"] = kpi_tbl_show["PV01 assets"].map(lambda x: f"{x:,.0f}")
        kpi_tbl_show["PV01 liabilities"] = kpi_tbl_show["PV01 liabilities"].map(lambda x: f"{x:,.0f}")
        kpi_tbl_show["Mac duration assets"] = kpi_tbl_show["Mac duration assets"].round(4)
        kpi_tbl_show["Mac duration liabilities"] = kpi_tbl_show["Mac duration liabilities"].round(4)
        st.dataframe(kpi_tbl_show, use_container_width=True, hide_index=True)

        st.markdown("##### Key rate duration by tenor (1 bp localized bump)")
        try:
            pricing_for_krd = _build_pricing_for_selected_scenario()
            key_tenors = np.array([float(b.tenor_years) for b in ordered_specs if float(b.tenor_years) > 1e-12], dtype=float)
            if key_tenors.size > 0:
                asm_krd = asm_vis if isinstance(asm_vis, sp.ALMAssumptions) else sp.ALMAssumptions(
                    allocation=sp.ALMAllocationSpec(buckets=base_spec.buckets, weights=ws),
                    rebalance_band=float(band_pct) / 100.0,
                    rebalance_frequency_months=int(freq_m),
                    reinvest_rule=reinvest,  # type: ignore[arg-type]
                    disinvest_rule=disinvest,  # type: ignore[arg-type]
                    rebalance_policy=rebalance_policy,  # type: ignore[arg-type]
                    borrowing_policy=borrow_policy,  # type: ignore[arg-type]
                    borrowing_rate_mode=borrow_rate_mode,  # type: ignore[arg-type]
                    borrowing_rate_tenor_years=float(borrow_rate_tenor),
                    borrowing_spread_annual=float(borrow_spread_bps) / 10000.0,
                    borrowing_rate_annual=float(borrow_rate_pct) / 100.0,
                    liquidity_near_liquid_years=float(near_liq_y),
                )
                a0 = float(aum0)
                base_cf = np.asarray(pricing_for_krd.expected_total_cashflows, dtype=float)
                l0 = float(np.sum(base_cf * yc.discount_factors(pricing_for_krd.times_years, spread=spr)))
                net0 = max(1e-9, a0 - l0)
                w_krd = np.asarray(asm_krd.allocation.weights, dtype=float)
                bond_tenors = np.array([float(b.tenor_years) for b in asm_krd.allocation.buckets[1:]], dtype=float)
                df0_bonds = yc.discount_factors(bond_tenors, spread=spr)
                target_mv_bonds = w_krd[1:] * a0
                bond_faces = np.where(df0_bonds > 1e-15, target_mv_bonds / df0_bonds, 0.0)
                rows: list[dict[str, float | str]] = []
                for kt in key_tenors:
                    yc_b = _key_rate_bump_curve(yc, key_tenor_years=float(kt), key_tenors_years=key_tenors, bump_bps=1.0)
                    dfb_bonds = yc_b.discount_factors(bond_tenors, spread=spr)
                    a_b = float(w_krd[0] * a0 + np.sum(bond_faces * dfb_bonds))
                    l_b = float(np.sum(base_cf * yc_b.discount_factors(pricing_for_krd.times_years, spread=spr)))
                    rows.append(
                        {
                            "Tenor": f"{kt:g}Y",
                            "Tenor years": float(kt),
                            "Assets KRD": -((a_b - a0) / (max(1e-9, a0) * 1e-4)),
                            "Liabilities KRD": -((l_b - l0) / (max(1e-9, l0) * 1e-4)),
                            "Surplus KRD": -(((a_b - l_b) - (a0 - l0)) / (net0 * 1e-4)),
                        }
                    )
                krd_df = pd.DataFrame(rows).sort_values("Tenor years")
                krd_long = krd_df.melt(
                    id_vars=["Tenor", "Tenor years"],
                    value_vars=["Assets KRD", "Liabilities KRD", "Surplus KRD"],
                    var_name="Series",
                    value_name="Key rate duration",
                )
                krd_bars_df = krd_long[krd_long["Series"].isin(["Assets KRD", "Liabilities KRD"])].copy()
                krd_surplus_df = krd_long[krd_long["Series"] == "Surplus KRD"].copy()
                tenor_order = krd_df["Tenor"].tolist()

                bars = (
                    alt.Chart(krd_bars_df)
                    .mark_bar()
                    .encode(
                        x=alt.X("Tenor:N", sort=tenor_order, title="Key tenor"),
                        y=alt.Y("Key rate duration:Q", title="Assets/Liabilities KRD (years)"),
                        color=alt.Color(
                            "Series:N",
                            title="Series",
                            sort=["Assets KRD", "Liabilities KRD"],
                            legend=alt.Legend(orient="top", direction="horizontal"),
                        ),
                        xOffset=alt.XOffset("Series:N"),
                        tooltip=[
                            alt.Tooltip("Tenor:N"),
                            alt.Tooltip("Series:N"),
                            alt.Tooltip("Key rate duration:Q", format=".4f"),
                        ],
                    )
                )

                surplus_line = (
                    alt.Chart(krd_surplus_df)
                    .mark_line(color="#d62728", strokeWidth=3, point=True)
                    .encode(
                        x=alt.X("Tenor:N", sort=tenor_order, title="Key tenor"),
                        y=alt.Y("Key rate duration:Q", title="Surplus KRD (years)"),
                        tooltip=[
                            alt.Tooltip("Tenor:N"),
                            alt.Tooltip("Series:N"),
                            alt.Tooltip("Key rate duration:Q", format=".4f"),
                        ],
                    )
                )

                st.altair_chart(
                    alt.layer(bars, surplus_line).resolve_scale(y="independent").properties(height=320),
                    use_container_width=True,
                )
                st.caption(
                    "Interpretation: Surplus KRD is the key-rate sensitivity of net surplus (assets minus liabilities), "
                    "normalized by current surplus. Because the denominator is surplus rather than total assets or liabilities, "
                    "Surplus KRD can be much larger in magnitude when surplus is small."
                )
            else:
                st.info("No positive tenors available for key rate duration chart.")
        except Exception as ex:
            st.info(f"Key rate duration chart unavailable for current inputs: {ex!r}")


def main() -> None:
    st.set_page_config(page_title="Pricing Demo", layout="wide")
    with st.sidebar:
        st.title("Pricing Demo")
        page = st.radio(
            "Section",
            options=SECTION_ORDER,
            format_func=lambda x: SECTION_LABELS[x],
        )
        st.divider()
        st.caption(f"Project root: `{ROOT}`")

        st.subheader("Diagnostics export")
        # Diagnostics should be fully self-contained for offline review/debugging.
        # Always include everything (no include/exclude toggles).
        include_full_paths = True
        include_alm_buckets = True
        if st.button("Prepare diagnostics JSON", type="secondary"):
            pricing_res = st.session_state.get("pricing_res")
            pricing_contract = st.session_state.get("pricing_contract")
            pricing_excel_context = st.session_state.get("pricing_excel_context") or {}
            alm_last = st.session_state.get("alm_last")
            alm_last_assumptions = st.session_state.get("alm_last_assumptions")
            alm_current_assumptions = st.session_state.get("alm_current_assumptions")
            alm_current_aum0 = st.session_state.get("alm_current_initial_asset_market_value")

            if pricing_res is None or pricing_contract is None:
                st.warning("Run Pricing Run first to populate diagnostics.")
            else:
                ctx_yc = pricing_excel_context.get("yield_curve")
                ctx_mort = pricing_excel_context.get("mortality")
                ctx_exp = pricing_excel_context.get("expenses")
                payload: dict[str, Any] = {
                    "exported_at_utc": _dt.datetime.utcnow().isoformat() + "Z",
                    "pricing_run_id": st.session_state.get("pricing_run_id"),
                    "pricing_meta": st.session_state.get("pricing_meta") or {},
                    "pricing_run_inputs": st.session_state.get("pricing_run_inputs") or {},
                    "pricing": _pricing_result_to_dict(
                        pricing_res,
                        pricing_contract,
                        include_full=include_full_paths,
                    ),
                    "pricing_inputs": {
                        "horizon_age": pricing_excel_context.get("horizon_age"),
                        "valuation_year": pricing_excel_context.get("valuation_year"),
                        "spread": pricing_excel_context.get("spread"),
                        "yield_curve": _yield_curve_to_dict(ctx_yc) if isinstance(ctx_yc, sp.YieldCurve) else None,
                        "mortality": _mortality_to_dict(ctx_mort) if ctx_mort is not None else None,
                        "expenses": (
                            {
                                "policy_expense_dollars": float(getattr(ctx_exp, "policy_expense_dollars", float("nan"))),
                                "premium_expense_rate": float(getattr(ctx_exp, "premium_expense_rate", float("nan"))),
                                "monthly_expense_dollars": float(getattr(ctx_exp, "monthly_expense_dollars", float("nan"))),
                            }
                            if isinstance(ctx_exp, sp.ExpenseAssumptions)
                            else None
                        ),
                        "yield_mode": pricing_excel_context.get("yield_mode"),
                        "mortality_mode": pricing_excel_context.get("mortality_mode"),
                        "expense_mode": pricing_excel_context.get("expense_mode"),
                        "expense_annual_inflation": pricing_excel_context.get("expense_annual_inflation"),
                    },
                    "alm": None,
                    "alm_current": None,
                    "what_if": None,
                }

                current_pricing_run_id = st.session_state.get("pricing_run_id")
                alm_run_id = st.session_state.get("alm_last_pricing_run_id")
                whatif_run_id = st.session_state.get("whatif_last_pricing_run_id")

                if isinstance(alm_last, sp.ALMResult) and alm_run_id == current_pricing_run_id:
                    payload["alm"] = _alm_result_to_dict(
                        alm_last,
                        alm_last_assumptions if isinstance(alm_last_assumptions, sp.ALMAssumptions) else None,
                        include_buckets=include_alm_buckets,
                        include_full=include_full_paths,
                    )

                if isinstance(alm_current_assumptions, sp.ALMAssumptions):
                    payload["alm_current"] = {
                        "initial_asset_market_value": float(alm_current_aum0) if alm_current_aum0 is not None else None,
                        "assumptions": _alm_assumptions_to_dict(alm_current_assumptions),
                    }

                what_if_shocked_res = st.session_state.get("whatif_last_shocked_res")
                what_if_base_res = st.session_state.get("whatif_last_base_res")
                what_if_alm_base = st.session_state.get("whatif_last_alm_base")
                what_if_alm_after = st.session_state.get("whatif_last_alm_after")
                what_if_baseline_mc = st.session_state.get("whatif_last_baseline_mc")
                what_if_shocked_mc = st.session_state.get("whatif_last_shocked_mc")
                what_if_shocked_curve = st.session_state.get("whatif_last_shocked_curve")
                what_if_shocked_mortality = st.session_state.get("whatif_last_shocked_mortality")
                what_if_alm_assumptions = st.session_state.get("whatif_last_alm_assumptions")
                what_if_params = st.session_state.get("whatif_last_params") or {}

                if (
                    whatif_run_id == current_pricing_run_id
                    and
                    what_if_shocked_res is not None
                    and what_if_base_res is not None
                    and what_if_baseline_mc is not None
                    and what_if_shocked_mc is not None
                ):
                    payload["what_if"] = _whatif_result_to_dict(
                        base_res=what_if_base_res,
                        shocked_res=what_if_shocked_res,
                        baseline_mc=what_if_baseline_mc,
                        shocked_mc=what_if_shocked_mc,
                        whatif_params={
                            **what_if_params,
                            "shocked_curve": _yield_curve_to_dict(what_if_shocked_curve)
                            if isinstance(what_if_shocked_curve, sp.YieldCurve)
                            else None,
                            "shocked_mortality": _mortality_to_dict(what_if_shocked_mortality)
                            if what_if_shocked_mortality is not None
                            else None,
                        },
                        alm_base=what_if_alm_base,
                        alm_after=what_if_alm_after,
                        asm=what_if_alm_assumptions if isinstance(what_if_alm_assumptions, sp.ALMAssumptions) else None,
                        include_full=include_full_paths,
                    )

                st.session_state["diagnostics_json_bytes"] = json.dumps(payload, default=str, ensure_ascii=False, indent=2).encode("utf-8")
                st.session_state["diagnostics_json_filename"] = (
                    f"pricing_diagnostics_{_dt.datetime.utcnow().strftime('%Y%m%d_%H%M%S')}.json"
                )
                st.success("Diagnostics JSON prepared. Use Download below.")

        diag_bytes = st.session_state.get("diagnostics_json_bytes")
        diag_name = st.session_state.get("diagnostics_json_filename") or "pricing_diagnostics.json"
        if isinstance(diag_bytes, (bytes, bytearray)) and diag_bytes:
            st.download_button(
                "Download diagnostics JSON",
                data=diag_bytes,
                file_name=diag_name,
                mime="application/json",
                type="primary",
            )

    if page == "overview":
        _render_overview()
    elif page == "run":
        _render_run_and_results()
    elif page == "alm":
        _render_alm_section()
    elif page == "what_if":
        _render_what_if_studio()
    elif page == "excel_replicator":
        _render_excel_replicator()
    else:
        render_unit_tests_page(embedded=True)


if __name__ == "__main__":
    main()
