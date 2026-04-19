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
3. Keys in :data:`PRICING_RUN_NUMBER_INPUT_KEYS` must not be ``setdefault``'d into
   ``st.session_state`` before the widget — that duplicates binding with ``value=``
   and triggers Streamlit warnings. The Pricing Run page stores coerced defaults in
   ``_pricing_run_numeric_seeds`` for first-paint seeding instead.
"""

from __future__ import annotations

from collections.abc import Mapping, MutableMapping, Sequence
from typing import Any

import streamlit as st

import pricing_projection as sp
from product_registry import (
    ProductType,
    get_product_default_mortality_mode,
    get_term_contract_ui_config,
    parse_term_length_label_to_years,
)

class RUN_KEY:
    """Canonical Streamlit ``st.session_state`` keys for the Pricing Run page.

    Why a namespace class instead of bare module constants?
    -------------------------------------------------------
    * Discoverability: an IDE auto-complete on ``RUN_KEY.<TAB>`` enumerates
      every legal key, so a typo (``RUN_KEY.ISSUE_AG``) becomes a static
      ``AttributeError`` instead of a silent ``st.session_state.get`` miss.
    * Single source of truth: :data:`RUN_STATE_KEY_NAMES` (the frozenset of
      string values) is derived from this class via reflection so adding a
      key in *one* place propagates to the ratchet test, the
      ``PRICING_RUN_NUMBER_INPUT_KEYS`` subset, and any downstream consumer
      that imports the namespace.
    * Migration ratchet: ``tests/test_run_state_key_drift.py`` walks every
      ``.py`` in the repo and compares per-file occurrences of these
      string literals against a committed baseline. New code MUST use
      the symbols below; legacy ``pricing_ui.py`` literals are baselined
      and the count is allowed to *decrease* over time (the
      ``ui/MIGRATION.md`` decomposition deletes them naturally).

    Adding a new Pricing Run widget? Add a class attribute below, add it
    to ``build_run_form_seed_defaults`` if it needs a default, and (if
    it's a ``st.number_input``) add the new symbol to
    ``PRICING_RUN_NUMBER_INPUT_KEYS``. The ratchet test will pick it up
    on the next run.
    """

    # Identity / common
    PRODUCT_TYPE = "run_product_type"
    ISSUE_AGE = "run_issue_age"
    SEX = "run_sex"

    # SPIA-specific
    SPIA_BENEFIT_ANNUAL = "run_spia_benefit_annual"

    # Term-specific
    TERM_BENEFIT_ANNUAL = "run_term_benefit_annual"
    TERM_MONTHLY_PREMIUM = "run_term_monthly_premium"
    TERM_LENGTH = "run_term_length"
    TERM_PREMIUM_MODE = "run_term_premium_mode"
    TERM_BENEFIT_TIMING = "run_term_benefit_timing"

    # RILA-specific
    RILA_PARTICIPATION = "run_rila_participation"
    RILA_CAP = "run_rila_cap"
    RILA_FLOOR = "run_rila_floor"
    RILA_RIDER_FEE = "run_rila_rider_fee"

    # MYGA
    MYGA_SINGLE_PREMIUM = "run_myga_single_premium"
    MYGA_DECLARED_RATE = "run_myga_declared_rate"
    MYGA_GUARANTEE_YEARS = "run_myga_guarantee_years"

    # FIA
    FIA_SINGLE_PREMIUM = "run_fia_single_premium"
    FIA_PARTICIPATION = "run_fia_participation"
    FIA_CAP = "run_fia_cap"
    FIA_FLOOR = "run_fia_floor"
    FIA_HORIZON_YEARS = "run_fia_horizon_years"

    # VA
    VA_SINGLE_PREMIUM = "run_va_single_premium"
    VA_ME_CHARGE = "run_va_me_charge"
    VA_HORIZON_YEARS = "run_va_horizon_years"

    # WL (Whole Life)
    WL_FACE_AMOUNT = "run_wl_face_amount"
    WL_SMOKER_CLASS = "run_wl_smoker_class"

    # UL
    UL_FACE_AMOUNT = "run_ul_face_amount"
    UL_SMOKER_CLASS = "run_ul_smoker_class"
    UL_SINGLE_PREMIUM = "run_ul_single_premium"
    UL_PREMIUM_LOAD = "run_ul_premium_load"
    UL_MONTHLY_EXPENSE = "run_ul_monthly_expense"
    UL_DECLARED_RATE = "run_ul_declared_rate"

    # IUL
    IUL_FACE_AMOUNT = "run_iul_face_amount"
    IUL_SMOKER_CLASS = "run_iul_smoker_class"
    IUL_SINGLE_PREMIUM = "run_iul_single_premium"
    IUL_PREMIUM_LOAD = "run_iul_premium_load"
    IUL_MONTHLY_EXPENSE = "run_iul_monthly_expense"
    IUL_PARTICIPATION = "run_iul_participation"
    IUL_CAP = "run_iul_cap"
    IUL_FLOOR = "run_iul_floor"

    # VUL
    VUL_FACE_AMOUNT = "run_vul_face_amount"
    VUL_SMOKER_CLASS = "run_vul_smoker_class"
    VUL_SINGLE_PREMIUM = "run_vul_single_premium"
    VUL_PREMIUM_LOAD = "run_vul_premium_load"
    VUL_MONTHLY_EXPENSE = "run_vul_monthly_expense"

    # Yield curve / discounting
    Y_MODE = "run_y_mode"
    FLAT_RATE = "run_flat_rate"
    ZERO_CSV = "run_zero_csv"
    PAR_CSV = "run_par_csv"
    COUPON_FREQ = "run_coupon_freq"
    SPREAD = "run_spread"

    # Mortality
    M_MODE = "run_m_mode"
    QX_CSV = "run_qx_csv"
    RP_XLSX = "run_rp_xlsx"
    RP_OUT = "run_rp_out"
    MP_XLSX = "run_mp_xlsx"
    MP_OUT = "run_mp_out"

    # Expenses
    EXPENSE_MODE = "run_expense_mode"
    EXPENSES_CSV = "run_expenses_csv"
    POLICY_EXPENSE = "run_policy_expense"
    PREMIUM_EXPENSE_PCT = "run_premium_expense_pct"
    MONTHLY_EXPENSE = "run_monthly_expense"
    EXPENSE_INFLATION_PCT = "run_expense_inflation_pct"

    # Horizon / valuation
    HORIZON_AGE = "run_horizon_age"
    VALUATION_YEAR = "run_valuation_year"

    # Index scenario
    USE_INDEX = "run_use_index"
    INDEX_CSV = "run_index_csv"

    # Monte Carlo
    MC_ENABLE = "run_mc_enable"
    MC_N_SIMS = "run_mc_n_sims"
    MC_SEED = "run_mc_seed"
    MC_DRIFT_PCT = "run_mc_drift_pct"
    MC_VOL_PCT = "run_mc_vol_pct"
    MC_S0 = "run_mc_s0"


def _all_run_key_names() -> frozenset[str]:
    """Reflectively enumerate every ``"run_..."`` constant on :class:`RUN_KEY`.

    Drives the ratchet test and any other consumer that needs the
    canonical set. We deliberately re-derive on every import (rather than
    caching) so adding a class attribute to :class:`RUN_KEY` is the
    *only* edit required.
    """
    names: set[str] = set()
    for attr in vars(RUN_KEY).values():
        if isinstance(attr, str) and attr.startswith("run_"):
            names.add(attr)
    return frozenset(names)


RUN_STATE_KEY_NAMES: frozenset[str] = _all_run_key_names()


# Keys managed by :func:`run_number_input`. Do not ``setdefault`` these on the Pricing Run page
# — doing so puts them in ``_new_session_state`` before the widget runs, and passing ``value=``
# to ``st.number_input`` triggers Streamlit's "default value + Session State API" warning.
PRICING_RUN_NUMBER_INPUT_KEYS: frozenset[str] = frozenset(
    {
        RUN_KEY.ISSUE_AGE,
        RUN_KEY.SPIA_BENEFIT_ANNUAL,
        RUN_KEY.TERM_BENEFIT_ANNUAL,
        RUN_KEY.TERM_MONTHLY_PREMIUM,
        RUN_KEY.FLAT_RATE,
        RUN_KEY.COUPON_FREQ,
        RUN_KEY.POLICY_EXPENSE,
        RUN_KEY.PREMIUM_EXPENSE_PCT,
        RUN_KEY.MONTHLY_EXPENSE,
        RUN_KEY.VALUATION_YEAR,
        RUN_KEY.HORIZON_AGE,
        RUN_KEY.SPREAD,
        RUN_KEY.EXPENSE_INFLATION_PCT,
        RUN_KEY.MC_N_SIMS,
        RUN_KEY.MC_SEED,
        RUN_KEY.MC_DRIFT_PCT,
        RUN_KEY.MC_VOL_PCT,
        RUN_KEY.MC_S0,
        RUN_KEY.RILA_PARTICIPATION,
        RUN_KEY.RILA_CAP,
        RUN_KEY.RILA_FLOOR,
        RUN_KEY.RILA_RIDER_FEE,
        # MYGA
        RUN_KEY.MYGA_SINGLE_PREMIUM,
        RUN_KEY.MYGA_DECLARED_RATE,
        RUN_KEY.MYGA_GUARANTEE_YEARS,
        # FIA
        RUN_KEY.FIA_SINGLE_PREMIUM,
        RUN_KEY.FIA_PARTICIPATION,
        RUN_KEY.FIA_CAP,
        RUN_KEY.FIA_FLOOR,
        RUN_KEY.FIA_HORIZON_YEARS,
        # VA
        RUN_KEY.VA_SINGLE_PREMIUM,
        RUN_KEY.VA_ME_CHARGE,
        RUN_KEY.VA_HORIZON_YEARS,
        # WL
        RUN_KEY.WL_FACE_AMOUNT,
        # UL
        RUN_KEY.UL_FACE_AMOUNT,
        RUN_KEY.UL_SINGLE_PREMIUM,
        RUN_KEY.UL_PREMIUM_LOAD,
        RUN_KEY.UL_MONTHLY_EXPENSE,
        RUN_KEY.UL_DECLARED_RATE,
        # IUL
        RUN_KEY.IUL_FACE_AMOUNT,
        RUN_KEY.IUL_SINGLE_PREMIUM,
        RUN_KEY.IUL_PREMIUM_LOAD,
        RUN_KEY.IUL_MONTHLY_EXPENSE,
        RUN_KEY.IUL_PARTICIPATION,
        RUN_KEY.IUL_CAP,
        RUN_KEY.IUL_FLOOR,
        # VUL
        RUN_KEY.VUL_FACE_AMOUNT,
        RUN_KEY.VUL_SINGLE_PREMIUM,
        RUN_KEY.VUL_PREMIUM_LOAD,
        RUN_KEY.VUL_MONTHLY_EXPENSE,
    }
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
    if raw is None:
        seeds = st.session_state.get("_pricing_run_numeric_seeds") or {}
        if key in seeds:
            raw = seeds[key]
    coerced = coerce_numeric_widget_value(
        raw,
        default,
        min_value=min_v,
        max_value=max_v,
        replace_non_positive=replace_non_positive,
    )
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
        RUN_KEY.PRODUCT_TYPE: product_default,
        RUN_KEY.ISSUE_AGE: int(saved_inputs.get("issue_age", 65)),
        RUN_KEY.SEX: str(saved_inputs.get("sex", "male")),
        RUN_KEY.TERM_MONTHLY_PREMIUM: seeded_term_monthly_premium,
        RUN_KEY.Y_MODE: str(meta.get("yield_mode", "par_bootstrap")),
        RUN_KEY.M_MODE: str(
            meta.get("mortality_mode", get_product_default_mortality_mode(default_product_type))
        ),
        RUN_KEY.EXPENSE_MODE: str(meta.get("expense_mode", "csv")),
        RUN_KEY.HORIZON_AGE: int(saved_inputs.get("horizon_age", 110)),
        RUN_KEY.VALUATION_YEAR: int(saved_inputs.get("valuation_year", 2025)),
        RUN_KEY.SPREAD: float(saved_inputs.get("spread", 0.0)),
        RUN_KEY.USE_INDEX: bool(saved_inputs.get("use_index", True)),
        RUN_KEY.INDEX_CSV: str(
            saved_inputs.get("index_scenario_csv_path") or sp.DEFAULT_SP500_SCENARIO_CSV
        ),
        RUN_KEY.EXPENSE_INFLATION_PCT: float(
            saved_inputs.get("expense_annual_inflation", 0.025) * 100.0
        ),
        RUN_KEY.MC_ENABLE: bool(saved_inputs.get("mc_enabled", True)),
        RUN_KEY.MC_N_SIMS: int(saved_inputs.get("mc_n_sims", 100)),
        RUN_KEY.MC_SEED: int(saved_inputs.get("mc_seed", 42)),
        RUN_KEY.MC_DRIFT_PCT: float(saved_inputs.get("mc_annual_drift", 0.06) * 100.0),
        RUN_KEY.MC_VOL_PCT: float(saved_inputs.get("mc_annual_vol", 0.15) * 100.0),
        RUN_KEY.MC_S0: float(saved_inputs.get("mc_s0", 100.0)),
        RUN_KEY.QX_CSV: _nonblank_str(
            saved_inputs, "mortality_qx_csv", sp.DEFAULT_MORTALITY_QX_CSV
        ),
        RUN_KEY.RP_XLSX: _nonblank_str(saved_inputs, "mortality_rp_xlsx", sp.DEFAULT_RP2014_XLSX),
        RUN_KEY.RP_OUT: _nonblank_str(
            saved_inputs, "mortality_rp_out_csv", sp.DEFAULT_RP2014_MALE_HEALTHY_QX_CSV
        ),
        RUN_KEY.MP_XLSX: _nonblank_str(saved_inputs, "mortality_mp_xlsx", sp.DEFAULT_MP2016_XLSX),
        RUN_KEY.MP_OUT: _nonblank_str(
            saved_inputs, "mortality_mp_out_csv", sp.DEFAULT_MP2016_MALE_IMPROVEMENT_CSV
        ),
        # Separate keys per product; fallbacks match historical expander defaults.
        RUN_KEY.SPIA_BENEFIT_ANNUAL: float(saved_inputs.get("benefit_annual", 100_000.0)),
        RUN_KEY.TERM_BENEFIT_ANNUAL: float(
            saved_inputs.get("benefit_annual", term_ui.default_death_benefit)
        ),
        RUN_KEY.TERM_LENGTH: str(
            saved_inputs.get("term_length", term_ui.term_length_options[0])
        ),
        RUN_KEY.TERM_PREMIUM_MODE: str(
            saved_inputs.get("term_premium_mode", term_ui.premium_mode_options[0])
        ),
        RUN_KEY.TERM_BENEFIT_TIMING: str(
            saved_inputs.get("term_benefit_timing", term_ui.benefit_timing_options[0])
        ),
        RUN_KEY.FLAT_RATE: 0.04,
        RUN_KEY.ZERO_CSV: sp.DEFAULT_ZERO_CURVE_CSV,
        RUN_KEY.PAR_CSV: sp.DEFAULT_PAR_CURVE_CSV,
        RUN_KEY.COUPON_FREQ: 2,
        RUN_KEY.EXPENSES_CSV: sp.DEFAULT_EXPENSES_CSV,
        RUN_KEY.POLICY_EXPENSE: 0.0,
        RUN_KEY.PREMIUM_EXPENSE_PCT: 0.0,
        RUN_KEY.MONTHLY_EXPENSE: 0.0,
        RUN_KEY.RILA_PARTICIPATION: float(saved_inputs.get("rila_participation", 1.0)),
        RUN_KEY.RILA_CAP: float(saved_inputs.get("rila_cap", 0.10)),
        RUN_KEY.RILA_FLOOR: float(saved_inputs.get("rila_floor", 0.0)),
        RUN_KEY.RILA_RIDER_FEE: float(saved_inputs.get("rila_rider_fee_annual", 0.01)),
        # MYGA
        RUN_KEY.MYGA_SINGLE_PREMIUM: float(saved_inputs.get("myga_single_premium", 100_000.0)),
        RUN_KEY.MYGA_DECLARED_RATE: float(saved_inputs.get("myga_declared_rate", 0.045)),
        RUN_KEY.MYGA_GUARANTEE_YEARS: int(saved_inputs.get("myga_guarantee_years", 5)),
        # FIA
        RUN_KEY.FIA_SINGLE_PREMIUM: float(saved_inputs.get("fia_single_premium", 100_000.0)),
        RUN_KEY.FIA_PARTICIPATION: float(saved_inputs.get("fia_participation", 0.80)),
        RUN_KEY.FIA_CAP: float(saved_inputs.get("fia_cap", 0.07)),
        RUN_KEY.FIA_FLOOR: float(saved_inputs.get("fia_floor", 0.0)),
        RUN_KEY.FIA_HORIZON_YEARS: int(saved_inputs.get("fia_horizon_years", 10)),
        # VA
        RUN_KEY.VA_SINGLE_PREMIUM: float(saved_inputs.get("va_single_premium", 100_000.0)),
        RUN_KEY.VA_ME_CHARGE: float(saved_inputs.get("va_me_charge_annual", 0.014)),
        RUN_KEY.VA_HORIZON_YEARS: int(saved_inputs.get("va_horizon_years", 20)),
        # WL
        RUN_KEY.WL_FACE_AMOUNT: float(saved_inputs.get("wl_face_amount", 250_000.0)),
        RUN_KEY.WL_SMOKER_CLASS: str(saved_inputs.get("wl_smoker_class", "nonsmoker")),
        # UL
        RUN_KEY.UL_FACE_AMOUNT: float(saved_inputs.get("ul_face_amount", 250_000.0)),
        RUN_KEY.UL_SMOKER_CLASS: str(saved_inputs.get("ul_smoker_class", "nonsmoker")),
        RUN_KEY.UL_SINGLE_PREMIUM: float(saved_inputs.get("ul_single_premium", 25_000.0)),
        RUN_KEY.UL_PREMIUM_LOAD: float(saved_inputs.get("ul_premium_load_pct", 0.06)),
        RUN_KEY.UL_MONTHLY_EXPENSE: float(saved_inputs.get("ul_monthly_expense_charge", 7.50)),
        RUN_KEY.UL_DECLARED_RATE: float(saved_inputs.get("ul_declared_rate_annual", 0.04)),
        # IUL
        RUN_KEY.IUL_FACE_AMOUNT: float(saved_inputs.get("iul_face_amount", 250_000.0)),
        RUN_KEY.IUL_SMOKER_CLASS: str(saved_inputs.get("iul_smoker_class", "nonsmoker")),
        RUN_KEY.IUL_SINGLE_PREMIUM: float(saved_inputs.get("iul_single_premium", 25_000.0)),
        RUN_KEY.IUL_PREMIUM_LOAD: float(saved_inputs.get("iul_premium_load_pct", 0.06)),
        RUN_KEY.IUL_MONTHLY_EXPENSE: float(saved_inputs.get("iul_monthly_expense_charge", 7.50)),
        RUN_KEY.IUL_PARTICIPATION: float(saved_inputs.get("iul_participation", 1.0)),
        RUN_KEY.IUL_CAP: float(saved_inputs.get("iul_cap", 0.10)),
        RUN_KEY.IUL_FLOOR: float(saved_inputs.get("iul_floor", 0.0)),
        # VUL
        RUN_KEY.VUL_FACE_AMOUNT: float(saved_inputs.get("vul_face_amount", 250_000.0)),
        RUN_KEY.VUL_SMOKER_CLASS: str(saved_inputs.get("vul_smoker_class", "nonsmoker")),
        RUN_KEY.VUL_SINGLE_PREMIUM: float(saved_inputs.get("vul_single_premium", 25_000.0)),
        RUN_KEY.VUL_PREMIUM_LOAD: float(saved_inputs.get("vul_premium_load_pct", 0.06)),
        RUN_KEY.VUL_MONTHLY_EXPENSE: float(saved_inputs.get("vul_monthly_expense_charge", 7.50)),
    }
    return defaults


# Columns for manual portfolio entry / assembled inforce DataFrame (CSV-shaped).
PORTFOLIO_INFORCE_SCRATCH_COLUMNS: tuple[str, ...] = (
    "policy_id",
    "product_type",
    "issue_age",
    "sex",
    "benefit_annual",
    "death_benefit",
    "monthly_premium",
    "term_years",
    "participation",
    "cap",
    "floor",
    "rider_fee_annual",
    "me_charge_annual",
    "single_premium",
    "declared_rate_annual",
    "guarantee_years",
    "face_amount",
    "horizon_years",
    "gmdb_basis",
)


def default_inforce_scratch_row(
    product_type: ProductType,
    *,
    preserve: Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    """Wide inforce-shaped row defaults aligned with :func:`build_run_form_seed_defaults`.

    Maps ``RUN_KEY`` / Term UI config into CSV column names used by
    :func:`inforce_parsers.contract_from_inforce_row`. Unused fields are ``None``
    so they become NaN in a DataFrame (same as blank CSV cells).

    *preserve* may include ``policy_id``, ``issue_age``, ``sex`` to keep across
    product switches (Portfolio manual entry UX).
    """
    seeds = build_run_form_seed_defaults(
        product_default=product_type.value,
        saved_inputs={},
        meta={},
        default_product_type=product_type,
    )
    row: dict[str, Any] = {c: None for c in PORTFOLIO_INFORCE_SCRATCH_COLUMNS}
    row["product_type"] = product_type.value
    row["issue_age"] = int(seeds[RUN_KEY.ISSUE_AGE])
    row["sex"] = str(seeds[RUN_KEY.SEX])
    row["policy_id"] = ""

    if product_type == ProductType.SPIA:
        row["benefit_annual"] = float(seeds[RUN_KEY.SPIA_BENEFIT_ANNUAL])
    elif product_type == ProductType.TERM_LIFE:
        row["death_benefit"] = float(seeds[RUN_KEY.TERM_BENEFIT_ANNUAL])
        row["monthly_premium"] = float(seeds[RUN_KEY.TERM_MONTHLY_PREMIUM])
        row["term_years"] = int(
            parse_term_length_label_to_years(str(seeds[RUN_KEY.TERM_LENGTH]))
        )
    elif product_type == ProductType.RILA:
        row["participation"] = float(seeds[RUN_KEY.RILA_PARTICIPATION])
        row["cap"] = float(seeds[RUN_KEY.RILA_CAP])
        row["floor"] = float(seeds[RUN_KEY.RILA_FLOOR])
        row["rider_fee_annual"] = float(seeds[RUN_KEY.RILA_RIDER_FEE])
    elif product_type == ProductType.MYGA:
        row["single_premium"] = float(seeds[RUN_KEY.MYGA_SINGLE_PREMIUM])
        row["declared_rate_annual"] = float(seeds[RUN_KEY.MYGA_DECLARED_RATE])
        row["guarantee_years"] = int(seeds[RUN_KEY.MYGA_GUARANTEE_YEARS])
    elif product_type == ProductType.FIA:
        row["single_premium"] = float(seeds[RUN_KEY.FIA_SINGLE_PREMIUM])
        row["participation"] = float(seeds[RUN_KEY.FIA_PARTICIPATION])
        row["cap"] = float(seeds[RUN_KEY.FIA_CAP])
        row["floor"] = float(seeds[RUN_KEY.FIA_FLOOR])
        row["rider_fee_annual"] = 0.0
        row["horizon_years"] = int(seeds[RUN_KEY.FIA_HORIZON_YEARS])
    elif product_type == ProductType.VARIABLE_ANNUITY:
        row["single_premium"] = float(seeds[RUN_KEY.VA_SINGLE_PREMIUM])
        row["me_charge_annual"] = float(seeds[RUN_KEY.VA_ME_CHARGE])
        row["horizon_years"] = int(seeds[RUN_KEY.VA_HORIZON_YEARS])
        row["gmdb_basis"] = "return_of_premium"
    elif product_type == ProductType.WHOLE_LIFE:
        row["face_amount"] = float(seeds[RUN_KEY.WL_FACE_AMOUNT])
    elif product_type == ProductType.UNIVERSAL_LIFE:
        row["face_amount"] = float(seeds[RUN_KEY.UL_FACE_AMOUNT])
        row["single_premium"] = float(seeds[RUN_KEY.UL_SINGLE_PREMIUM])
    elif product_type == ProductType.INDEXED_UL:
        row["face_amount"] = float(seeds[RUN_KEY.IUL_FACE_AMOUNT])
        row["single_premium"] = float(seeds[RUN_KEY.IUL_SINGLE_PREMIUM])
        row["participation"] = float(seeds[RUN_KEY.IUL_PARTICIPATION])
        row["cap"] = float(seeds[RUN_KEY.IUL_CAP])
        row["floor"] = float(seeds[RUN_KEY.IUL_FLOOR])
    elif product_type == ProductType.VARIABLE_UL:
        row["face_amount"] = float(seeds[RUN_KEY.VUL_FACE_AMOUNT])
        row["single_premium"] = float(seeds[RUN_KEY.VUL_SINGLE_PREMIUM])
    else:
        raise NotImplementedError(f"portfolio defaults not wired for {product_type!r}")

    if preserve:
        for k in ("policy_id", "issue_age", "sex"):
            if k not in preserve:
                continue
            v = preserve[k]
            if v is None or (isinstance(v, float) and str(v) == "nan"):
                continue
            if k == "issue_age":
                try:
                    row["issue_age"] = int(v)
                except (TypeError, ValueError):
                    pass
            elif k == "sex":
                s = str(v).strip().lower()
                if s in ("male", "female"):
                    row["sex"] = s
            elif k == "policy_id":
                row["policy_id"] = str(v).strip()
    return row


class PORTFOLIO_KEY:
    """Session-state keys for the Portfolio (multi-policy) section (not ``run_*``)."""

    RESULT = "portfolio_res"
    RUN_ID = "portfolio_run_id"
    UPLOAD_NAME = "portfolio_upload_name"
    MANUAL_ROWS = "portfolio_manual_row_ids"
    # Sidebar-only: show Portfolio section when ``portfolio_v1_enabled()`` is False.
    UI_FORCE_SIDEBAR = "_pricing_ui_show_portfolio_sidebar"
