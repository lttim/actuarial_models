"""Assumption provenance views for pricing and portfolio demo runs."""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

from data_registry import REGISTRY, DataArtifact

ROOT = Path(__file__).resolve().parent
APPROVALS_PATH = ROOT / "data" / "assumptions" / "assumption_approvals.json"
_BY_NAME: dict[str, DataArtifact] = {a.name: a for a in REGISTRY}


@dataclass(frozen=True, slots=True)
class AssumptionApproval:
    artifact_name: str
    assumption_family: str
    approval_id: str
    approved_by: str
    challenged_by: str
    approval_date: str
    valid_from: str
    valid_to: str
    intended_use: str
    status: str
    requires_waiver_for_release: bool
    notes: str


def _load_approvals() -> dict[str, AssumptionApproval]:
    if not APPROVALS_PATH.is_file():
        return {}
    raw = json.loads(APPROVALS_PATH.read_text(encoding="utf-8"))
    return {
        str(row["artifact_name"]): AssumptionApproval(**row)
        for row in raw
        if isinstance(row, dict) and "artifact_name" in row
    }


def _artifact_by_path(path: str | None) -> DataArtifact | None:
    if not path:
        return None
    try:
        rp = Path(path).resolve()
    except (OSError, RuntimeError):
        return None
    for artifact in REGISTRY:
        try:
            if artifact.path.resolve() == rp:
                return artifact
        except (OSError, RuntimeError):
            continue
    return None


def _entry(
    *,
    role: str,
    artifact: DataArtifact | None,
    path: str | None,
    mode: str,
    approvals: dict[str, AssumptionApproval],
) -> dict[str, Any]:
    approval = approvals.get(artifact.name) if artifact is not None else None
    source = artifact.source if artifact is not None else "User-supplied or generated in-session."
    warning = ""
    if "PLACEHOLDER" in source.upper() or "SYNTHETIC" in source.upper():
        warning = "Placeholder/synthetic assumption: not production-approved without waiver."
    out = {
        "role": role,
        "mode": mode,
        "artifact_name": artifact.name if artifact is not None else "(custom)",
        "version": artifact.version if artifact is not None else "",
        "path": str(artifact.path) if artifact is not None else (path or ""),
        "sha256": artifact.sha256 if artifact is not None else "",
        "source": source,
        "approval_id": approval.approval_id if approval is not None else "",
        "approved_by": approval.approved_by if approval is not None else "",
        "challenged_by": approval.challenged_by if approval is not None else "",
        "status": approval.status if approval is not None else "unregistered",
        "intended_use": approval.intended_use if approval is not None else "user_supplied",
        "valid_to": approval.valid_to if approval is not None else "",
        "requires_waiver_for_release": (
            approval.requires_waiver_for_release if approval is not None else False
        ),
        "warning": warning,
    }
    return out


def provenance_rows_from_pricing_state(
    *,
    pricing_meta: dict[str, Any],
    pricing_run_inputs: dict[str, Any],
    pricing_excel_context: dict[str, Any],
) -> list[dict[str, Any]]:
    """Return assumption artifact rows for the active pricing run."""
    approvals = _load_approvals()
    rows: list[dict[str, Any]] = []

    y_mode = str(pricing_meta.get("yield_mode") or pricing_excel_context.get("yield_mode") or "")
    if y_mode == "par_bootstrap":
        for path, role in [
            (pricing_run_inputs.get("yield_par_csv"), "Yield curve par source"),
            (None, "Yield curve bootstrapped zero"),
        ]:
            artifact = (
                _artifact_by_path(path)
                if path
                else next((a for a in REGISTRY if a.name == "treasury_zero_curve"), None)
            )
            rows.append(
                _entry(role=role, artifact=artifact, path=path, mode=y_mode, approvals=approvals)
            )
    elif y_mode == "zero_csv":
        path = pricing_run_inputs.get("yield_zero_csv")
        rows.append(
            _entry(
                role="Yield curve zero source",
                artifact=_artifact_by_path(path),
                path=path,
                mode=y_mode,
                approvals=approvals,
            )
        )
    else:
        rows.append(
            _entry(role="Yield curve", artifact=None, path=None, mode=y_mode, approvals=approvals)
        )

    m_mode = str(
        pricing_meta.get("mortality_mode") or pricing_excel_context.get("mortality_mode") or ""
    )
    mortality_paths = [
        pricing_run_inputs.get("mortality_qx_csv"),
        pricing_run_inputs.get("mortality_rp_out_csv"),
        pricing_run_inputs.get("mortality_mp_out_csv"),
    ]
    if m_mode == "cso_2017_ult":
        for artifact in REGISTRY:
            if artifact.version == "cso_2017_ult":
                rows.append(
                    _entry(
                        role="Mortality table",
                        artifact=artifact,
                        path=str(artifact.path),
                        mode=m_mode,
                        approvals=approvals,
                    )
                )
    elif m_mode in {"rp2014_mp2016", "qx_csv"}:
        if m_mode == "rp2014_mp2016" and not any(mortality_paths):
            mortality_paths = [
                str(_BY_NAME["rp2014_male_healthy_annuitant_qx"].path),
                str(_BY_NAME["mp2016_male_improvement_rates"].path),
            ]
        for path in mortality_paths:
            artifact = _artifact_by_path(path)
            if artifact is not None:
                rows.append(
                    _entry(
                        role="Mortality table",
                        artifact=artifact,
                        path=path,
                        mode=m_mode,
                        approvals=approvals,
                    )
                )
    else:
        rows.append(
            _entry(
                role="Mortality table", artifact=None, path=None, mode=m_mode, approvals=approvals
            )
        )

    expense_mode = str(
        pricing_meta.get("expense_mode") or pricing_excel_context.get("expense_mode") or ""
    )
    expense_path = pricing_run_inputs.get("expenses_csv_path")
    if expense_mode == "csv" and not expense_path:
        expense_path = str(_BY_NAME["expenses_assumptions_us_placeholders"].path)
    rows.append(
        _entry(
            role="Expenses",
            artifact=_artifact_by_path(expense_path),
            path=expense_path,
            mode=expense_mode,
            approvals=approvals,
        )
    )

    index_path = pricing_meta.get("index_scenario_csv_path") or pricing_run_inputs.get(
        "index_scenario_csv_path"
    )
    if index_path:
        rows.append(
            _entry(
                role="Index scenario",
                artifact=_artifact_by_path(index_path),
                path=index_path,
                mode="index_csv",
                approvals=approvals,
            )
        )

    return rows


def approvals_as_dicts() -> list[dict[str, Any]]:
    return [asdict(a) for a in _load_approvals().values()]


__all__ = [
    "AssumptionApproval",
    "approvals_as_dicts",
    "provenance_rows_from_pricing_state",
]
