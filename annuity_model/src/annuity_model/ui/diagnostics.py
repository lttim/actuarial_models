"""Diagnostics export helpers for the Streamlit sidebar."""

from __future__ import annotations

import datetime as dt
import json
from collections.abc import Callable, MutableMapping
from dataclasses import dataclass
from typing import Any


class MissingDiagnosticsInputError(RuntimeError):
    """Raised when no pricing run is available for diagnostics export."""


@dataclass(frozen=True, slots=True)
class DiagnosticsBuilders:
    """Callbacks supplied by the pricing app for engine-specific serialization."""

    active_provenance_rows: Callable[[], list[dict[str, Any]]]
    pricing_result_to_dict: Callable[..., dict[str, Any]]
    yield_curve_to_dict: Callable[[Any], dict[str, Any]]
    mortality_to_dict: Callable[[Any], dict[str, Any]]
    alm_result_to_dict: Callable[..., dict[str, Any]]
    alm_assumptions_to_dict: Callable[[Any], dict[str, Any]]
    whatif_result_to_dict: Callable[..., dict[str, Any]]
    is_yield_curve: Callable[[Any], bool]
    is_expense_assumptions: Callable[[Any], bool]
    is_alm_result: Callable[[Any], bool]
    is_alm_assumptions: Callable[[Any], bool]


def _expense_assumptions_to_dict(expense: Any) -> dict[str, float]:
    return {
        "policy_expense_dollars": float(getattr(expense, "policy_expense_dollars", float("nan"))),
        "premium_expense_rate": float(getattr(expense, "premium_expense_rate", float("nan"))),
        "monthly_expense_dollars": float(getattr(expense, "monthly_expense_dollars", float("nan"))),
    }


def _utc_now_naive() -> dt.datetime:
    return dt.datetime.now(dt.UTC).replace(tzinfo=None)


def build_diagnostics_payload(
    state: MutableMapping[str, Any],
    *,
    builders: DiagnosticsBuilders,
    exported_at_utc: dt.datetime | None = None,
    include_full_paths: bool = True,
    include_alm_buckets: bool = True,
) -> dict[str, Any]:
    """Build the self-contained diagnostics payload from Streamlit state."""
    pricing_res = state.get("pricing_res")
    pricing_contract = state.get("pricing_contract")
    if pricing_res is None or pricing_contract is None:
        raise MissingDiagnosticsInputError("Run Pricing Run first to populate diagnostics.")

    pricing_excel_context = state.get("pricing_excel_context") or {}
    ctx_yc = pricing_excel_context.get("yield_curve")
    ctx_mort = pricing_excel_context.get("mortality")
    ctx_exp = pricing_excel_context.get("expenses")
    current_pricing_run_id = state.get("pricing_run_id")

    payload: dict[str, Any] = {
        "exported_at_utc": (exported_at_utc or _utc_now_naive()).isoformat() + "Z",
        "pricing_run_id": current_pricing_run_id,
        "pricing_meta": state.get("pricing_meta") or {},
        "pricing_run_inputs": state.get("pricing_run_inputs") or {},
        "pricing_run_summary": state.get("pricing_run_summary") or {},
        "run_ledger": {
            "path": state.get("pricing_run_ledger_path"),
            "error": state.get("pricing_run_ledger_error"),
            "record": state.get("pricing_run_ledger_record"),
        },
        "assumption_provenance": builders.active_provenance_rows(),
        "pricing": builders.pricing_result_to_dict(
            pricing_res,
            pricing_contract,
            include_full=include_full_paths,
        ),
        "pricing_inputs": {
            "horizon_age": pricing_excel_context.get("horizon_age"),
            "valuation_year": pricing_excel_context.get("valuation_year"),
            "spread": pricing_excel_context.get("spread"),
            "yield_curve": (
                builders.yield_curve_to_dict(ctx_yc) if builders.is_yield_curve(ctx_yc) else None
            ),
            "mortality": builders.mortality_to_dict(ctx_mort) if ctx_mort is not None else None,
            "expenses": (
                _expense_assumptions_to_dict(ctx_exp)
                if builders.is_expense_assumptions(ctx_exp)
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

    alm_last = state.get("alm_last")
    alm_last_assumptions = state.get("alm_last_assumptions")
    alm_run_id = state.get("alm_last_pricing_run_id")
    if builders.is_alm_result(alm_last) and alm_run_id == current_pricing_run_id:
        payload["alm"] = builders.alm_result_to_dict(
            alm_last,
            (alm_last_assumptions if builders.is_alm_assumptions(alm_last_assumptions) else None),
            include_buckets=include_alm_buckets,
            include_full=include_full_paths,
        )

    alm_current_assumptions = state.get("alm_current_assumptions")
    if builders.is_alm_assumptions(alm_current_assumptions):
        alm_current_aum0 = state.get("alm_current_initial_asset_market_value")
        payload["alm_current"] = {
            "initial_asset_market_value": (
                float(alm_current_aum0) if alm_current_aum0 is not None else None
            ),
            "assumptions": builders.alm_assumptions_to_dict(alm_current_assumptions),
        }

    _populate_what_if_payload(
        payload,
        state=state,
        builders=builders,
        current_pricing_run_id=current_pricing_run_id,
        include_full_paths=include_full_paths,
    )
    return payload


def _populate_what_if_payload(
    payload: dict[str, Any],
    *,
    state: MutableMapping[str, Any],
    builders: DiagnosticsBuilders,
    current_pricing_run_id: Any,
    include_full_paths: bool,
) -> None:
    whatif_run_id = state.get("whatif_last_pricing_run_id")
    what_if_shocked_res = state.get("whatif_last_shocked_res")
    what_if_base_res = state.get("whatif_last_base_res")
    what_if_baseline_mc = state.get("whatif_last_baseline_mc")
    what_if_shocked_mc = state.get("whatif_last_shocked_mc")

    pricing_meta_whatif = state.get("pricing_meta") or {}
    pt_whatif = str(pricing_meta_whatif.get("product_type", "spia"))
    what_if_need_mc = pt_whatif != "term_life"
    if (
        whatif_run_id != current_pricing_run_id
        or what_if_shocked_res is None
        or what_if_base_res is None
        or (what_if_need_mc and (what_if_baseline_mc is None or what_if_shocked_mc is None))
    ):
        return

    what_if_shocked_curve = state.get("whatif_last_shocked_curve")
    what_if_shocked_mortality = state.get("whatif_last_shocked_mortality")
    what_if_alm_assumptions = state.get("whatif_last_alm_assumptions")
    what_if_params = state.get("whatif_last_params") or {}
    payload["what_if"] = builders.whatif_result_to_dict(
        base_res=what_if_base_res,
        shocked_res=what_if_shocked_res,
        baseline_mc=what_if_baseline_mc,
        shocked_mc=what_if_shocked_mc,
        whatif_params={
            **what_if_params,
            "shocked_curve": (
                builders.yield_curve_to_dict(what_if_shocked_curve)
                if builders.is_yield_curve(what_if_shocked_curve)
                else None
            ),
            "shocked_mortality": (
                builders.mortality_to_dict(what_if_shocked_mortality)
                if what_if_shocked_mortality is not None
                else None
            ),
        },
        alm_base=state.get("whatif_last_alm_base"),
        alm_after=state.get("whatif_last_alm_after"),
        asm=(
            what_if_alm_assumptions
            if builders.is_alm_assumptions(what_if_alm_assumptions)
            else None
        ),
        include_full=include_full_paths,
    )


def render_diagnostics_export_sidebar(
    st_mod: Any,
    *,
    session_state: MutableMapping[str, Any],
    builders: DiagnosticsBuilders,
) -> None:
    """Render the diagnostics sidebar controls and store prepared JSON bytes."""
    st_mod.subheader("Diagnostics export")
    if st_mod.button("Prepare diagnostics JSON", type="secondary"):
        prepared_at = _utc_now_naive()
        try:
            payload = build_diagnostics_payload(
                session_state,
                builders=builders,
                exported_at_utc=prepared_at,
            )
        except MissingDiagnosticsInputError as exc:
            st_mod.warning(str(exc))
        else:
            session_state["diagnostics_json_bytes"] = json.dumps(
                payload, default=str, ensure_ascii=False, indent=2
            ).encode("utf-8")
            session_state["diagnostics_json_filename"] = (
                f"pricing_diagnostics_{prepared_at.strftime('%Y%m%d_%H%M%S')}.json"
            )
            st_mod.success("Diagnostics JSON prepared. Use Download below.")

    diag_bytes = session_state.get("diagnostics_json_bytes")
    diag_name = session_state.get("diagnostics_json_filename") or "pricing_diagnostics.json"
    if isinstance(diag_bytes, (bytes, bytearray)) and diag_bytes:
        st_mod.download_button(
            "Download diagnostics JSON",
            data=diag_bytes,
            file_name=diag_name,
            mime="application/json",
            type="primary",
        )
