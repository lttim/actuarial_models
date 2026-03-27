"""
Pricing Run tab: seeded session defaults and Streamlit numeric widget binding.

Problem
-------
``st.number_input`` uses ``value="min"`` by default, which resolves to ``min_value``
on first widget registration. The UI can then show **min_value** (valuation year 1950,
horizon age 1, issue age 0, SPIA benefit 0, …) even though seeded state uses larger
defaults. After "Run pricing", values look correct because reruns resync.

Mitigation
----------
1. :func:`build_run_form_seed_defaults` is the **only** place that defines initial
   ``run_*`` keys for the Pricing Run page (add new keys here when adding inputs).
2. Use :func:`run_number_input` for every keyed Pricing Run ``st.number_input`` so
   ``value=`` always matches coerced session state.
"""

from __future__ import annotations

from typing import Any, Mapping, MutableMapping, Sequence

import streamlit as st

import pricing_projection as sp
from product_registry import (
    ProductType,
    get_product_default_mortality_mode,
    get_term_contract_ui_config,
)


def coerce_numeric_widget_value(
    raw: Any,
    default: int | float,
    *,
    min_value: int | float | None = None,
    max_value: int | float | None = None,
    replace_non_positive: bool = False,
) -> int | float:
    """Parse *raw*; fall back to *default*, clamp to bounds, optional non-positive reset."""
    want_int = type(default) is int
    try:
        if raw is None:
            v: int | float = default
        else:
            v = int(raw) if want_int else float(raw)
    except (TypeError, ValueError):
        v = default
    fv = float(v)
    if replace_non_positive and fv <= 0.0:
        v = default
        fv = float(v)
    if min_value is not None and fv < float(min_value):
        v = int(min_value) if want_int and type(min_value) is int else float(min_value)
        fv = float(v)
    if max_value is not None and fv > float(max_value):
        v = int(max_value) if want_int and type(max_value) is int else float(max_value)
    return v


def run_number_input(
    label: str,
    key: str,
    *,
    default: int | float,
    **kwargs: Any,
) -> int | float:
    """
    Like ``st.number_input``, but always passes ``value=`` from coerced session state.

    Use ``replace_non_positive=True`` when 0 must mean "use default" (e.g. Term premium).
    """
    if "value" in kwargs:
        raise TypeError("run_number_input does not accept value=; use default= instead")
    replace_non_positive = bool(kwargs.pop("replace_non_positive", False))
    min_v = kwargs.get("min_value")
    max_v = kwargs.get("max_value")
    raw = st.session_state.get(key)
    coerced = coerce_numeric_widget_value(
        raw,
        default,
        min_value=min_v,
        max_value=max_v,
        replace_non_positive=replace_non_positive,
    )
    st.session_state[key] = coerced
    return st.number_input(label, value=coerced, key=key, **kwargs)


def ensure_session_choice(
    state: MutableMapping[str, Any],
    key: str,
    allowed: Sequence[str],
    default: str,
) -> None:
    """
    If *state*[*key*] is missing or not in *allowed*, set it to *default*.

    Use immediately before ``st.radio`` / ``st.selectbox`` when *allowed* changes by product
    so the widget never binds to a stale option (e.g. Term SSA mode lingering on SPIA).
    """
    cur = state.get(key)
    if cur is None or str(cur) not in allowed:
        state[key] = default


def _nonblank_str(saved: Mapping[str, Any], saved_key: str, fallback: str) -> str:
    raw = saved.get(saved_key, fallback)
    txt = str(raw) if raw is not None else ""
    return txt if txt.strip() else fallback


def build_run_form_seed_defaults(
    *,
    product_default: str,
    saved_inputs: Mapping[str, Any],
    meta: Mapping[str, Any],
    default_product_type: ProductType,
) -> dict[str, Any]:
    """Initial ``run_*`` keys for ``session_state.setdefault`` (single source of truth)."""
    term_ui = get_term_contract_ui_config()
    term_ui_default_monthly_premium = float(term_ui.default_monthly_premium)
    seeded_term_monthly_premium = float(
        saved_inputs.get("term_monthly_premium", term_ui_default_monthly_premium)
    )
    if seeded_term_monthly_premium <= 0.0:
        seeded_term_monthly_premium = term_ui_default_monthly_premium

    defaults: dict[str, Any] = {
        "run_product_type": product_default,
        "run_issue_age": int(saved_inputs.get("issue_age", 65)),
        "run_sex": str(saved_inputs.get("sex", "male")),
        "run_term_monthly_premium": seeded_term_monthly_premium,
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
        "run_qx_csv": _nonblank_str(saved_inputs, "mortality_qx_csv", sp.DEFAULT_MORTALITY_QX_CSV),
        "run_rp_xlsx": _nonblank_str(saved_inputs, "mortality_rp_xlsx", sp.DEFAULT_RP2014_XLSX),
        "run_rp_out": _nonblank_str(saved_inputs, "mortality_rp_out_csv", sp.DEFAULT_RP2014_MALE_HEALTHY_QX_CSV),
        "run_mp_xlsx": _nonblank_str(saved_inputs, "mortality_mp_xlsx", sp.DEFAULT_MP2016_XLSX),
        "run_mp_out": _nonblank_str(saved_inputs, "mortality_mp_out_csv", sp.DEFAULT_MP2016_MALE_IMPROVEMENT_CSV),
        # Separate keys per product; fallbacks match historical expander defaults.
        "run_spia_benefit_annual": float(saved_inputs.get("benefit_annual", 100_000.0)),
        "run_term_benefit_annual": float(
            saved_inputs.get("benefit_annual", term_ui.default_death_benefit)
        ),
        "run_term_length": str(saved_inputs.get("term_length", term_ui.term_length_options[0])),
        "run_term_premium_mode": str(saved_inputs.get("term_premium_mode", term_ui.premium_mode_options[0])),
        "run_term_benefit_timing": str(
            saved_inputs.get("term_benefit_timing", term_ui.benefit_timing_options[0])
        ),
        "run_flat_rate": 0.04,
        "run_zero_csv": sp.DEFAULT_ZERO_CURVE_CSV,
        "run_par_csv": sp.DEFAULT_PAR_CURVE_CSV,
        "run_coupon_freq": 2,
        "run_expenses_csv": sp.DEFAULT_EXPENSES_CSV,
        "run_policy_expense": 0.0,
        "run_premium_expense_pct": 0.0,
        "run_monthly_expense": 0.0,
    }
    return defaults
