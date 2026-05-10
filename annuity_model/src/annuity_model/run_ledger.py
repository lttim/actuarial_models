"""Durable pricing-run summary and ledger helpers.

``pricing_run_summary`` remains the stable payload constructor used by the UI
and diagnostics export. ``SQLiteRunLedger`` stores those payloads durably for
audit/replay and can export the complete ledger as JSON without changing the
existing public summary shape.
"""

from __future__ import annotations

import datetime as _dt
import hashlib
import json
import sqlite3
from collections.abc import Mapping
from pathlib import Path
from typing import Any

SCHEMA_VERSION = 1


def stable_input_hash(payload: Mapping[str, Any]) -> str:
    raw = json.dumps(payload, sort_keys=True, default=str, separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(raw).hexdigest()


def pricing_run_summary(
    *,
    run_id: str,
    product: str,
    scenario_id: str,
    assumption_artifacts: list[dict[str, Any]],
    input_payload: Mapping[str, Any],
    output_metrics: Mapping[str, Any],
    parity_status: str,
    created_at_utc: str | None = None,
    output_paths: list[str] | None = None,
    validation_status: str | None = None,
    waiver_status: str | None = None,
    assumption_evidence: Mapping[str, Any] | None = None,
    model_version: str | None = None,
    git_commit: str | None = None,
    metadata: Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    summary: dict[str, Any] = {
        "run_id": run_id,
        "product": product,
        "scenario_id": scenario_id,
        "assumption_artifacts": assumption_artifacts,
        "input_hash": stable_input_hash(input_payload),
        "output_metrics": dict(output_metrics),
        "parity_status": parity_status,
        "created_at": created_at_utc
        or _dt.datetime.now(_dt.UTC).isoformat(timespec="seconds").replace("+00:00", "Z"),
    }
    if output_paths is not None:
        summary["output_paths"] = list(output_paths)
    if validation_status is not None:
        summary["validation_status"] = validation_status
    if waiver_status is not None:
        summary["waiver_status"] = waiver_status
    if assumption_evidence is not None:
        summary["assumption_evidence"] = dict(assumption_evidence)
    if model_version is not None:
        summary["model_version"] = model_version
    if git_commit is not None:
        summary["git_commit"] = git_commit
    if metadata is not None:
        summary["metadata"] = dict(metadata)
    return summary


def default_ledger_path(base_dir: str | Path | None = None) -> Path:
    """Return the default SQLite ledger path under *base_dir* or cwd.

    The helper is intentionally deterministic and side-effect free; callers
    decide whether to create parent directories by instantiating
    :class:`SQLiteRunLedger`.
    """
    root = Path(base_dir) if base_dir is not None else Path.cwd()
    return root / "outputs" / "run_ledger.sqlite3"


class SQLiteRunLedger:
    """SQLite-backed pricing run ledger.

    The ledger stores the stable summary payload verbatim in ``summary_json``
    and indexes common audit fields in typed columns. ``record_pricing_run``
    uses ``INSERT OR REPLACE`` so rerunning a deterministic demo with the same
    run id updates the durable record instead of creating duplicate evidence.
    """

    def __init__(self, path: str | Path) -> None:
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.initialize()

    def initialize(self) -> None:
        with self._connect() as conn:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS pricing_runs (
                    run_id TEXT PRIMARY KEY,
                    product TEXT NOT NULL,
                    scenario_id TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    input_hash TEXT NOT NULL,
                    parity_status TEXT NOT NULL,
                    assumption_artifacts_json TEXT NOT NULL,
                    output_metrics_json TEXT NOT NULL,
                    output_paths_json TEXT NOT NULL,
                    validation_status TEXT,
                    waiver_status TEXT,
                    model_version TEXT,
                    git_commit TEXT,
                    metadata_json TEXT NOT NULL,
                    summary_json TEXT NOT NULL,
                    schema_version INTEGER NOT NULL
                )
                """
            )
            conn.execute(
                """
                CREATE INDEX IF NOT EXISTS idx_pricing_runs_product_created
                ON pricing_runs(product, created_at)
                """
            )
            conn.execute(
                """
                CREATE INDEX IF NOT EXISTS idx_pricing_runs_scenario_created
                ON pricing_runs(scenario_id, created_at)
                """
            )

    def record_pricing_run(self, summary: Mapping[str, Any]) -> dict[str, Any]:
        """Persist *summary* and return the normalized stored payload."""
        normalized = _normalize_summary(summary)
        with self._connect() as conn:
            conn.execute(
                """
                INSERT OR REPLACE INTO pricing_runs (
                    run_id,
                    product,
                    scenario_id,
                    created_at,
                    input_hash,
                    parity_status,
                    assumption_artifacts_json,
                    output_metrics_json,
                    output_paths_json,
                    validation_status,
                    waiver_status,
                    model_version,
                    git_commit,
                    metadata_json,
                    summary_json,
                    schema_version
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    normalized["run_id"],
                    normalized["product"],
                    normalized["scenario_id"],
                    normalized["created_at"],
                    normalized["input_hash"],
                    normalized["parity_status"],
                    _json_dumps(normalized["assumption_artifacts"]),
                    _json_dumps(normalized["output_metrics"]),
                    _json_dumps(normalized.get("output_paths", [])),
                    normalized.get("validation_status"),
                    normalized.get("waiver_status"),
                    normalized.get("model_version"),
                    normalized.get("git_commit"),
                    _json_dumps(normalized.get("metadata", {})),
                    _json_dumps(normalized),
                    SCHEMA_VERSION,
                ),
            )
        return normalized

    def get_pricing_run(self, run_id: str) -> dict[str, Any] | None:
        """Return one stored summary by run id, or ``None`` if absent."""
        with self._connect() as conn:
            row = conn.execute(
                "SELECT summary_json FROM pricing_runs WHERE run_id = ?",
                (run_id,),
            ).fetchone()
        if row is None:
            return None
        return dict(json.loads(row["summary_json"]))

    def iter_pricing_runs(
        self,
        *,
        product: str | None = None,
        scenario_id: str | None = None,
    ) -> list[dict[str, Any]]:
        """Return stored summaries in created-at order, optionally filtered."""
        query = "SELECT summary_json FROM pricing_runs ORDER BY created_at, run_id"
        params: tuple[str, ...] = ()
        if scenario_id is not None:
            query = (
                "SELECT summary_json FROM pricing_runs "
                "WHERE scenario_id = ? ORDER BY created_at, run_id"
            )
            params = (scenario_id,)
        if product is not None:
            query = (
                "SELECT summary_json FROM pricing_runs "
                "WHERE product = ? ORDER BY created_at, run_id"
            )
            params = (product,)
        if product is not None and scenario_id is not None:
            query = (
                "SELECT summary_json FROM pricing_runs "
                "WHERE product = ? AND scenario_id = ? ORDER BY created_at, run_id"
            )
            params = (product, scenario_id)
        with self._connect() as conn:
            rows = conn.execute(query, params).fetchall()
        return [dict(json.loads(row["summary_json"])) for row in rows]

    def export_json(self, out_path: str | Path | None = None) -> list[dict[str, Any]]:
        """Return all stored summaries and optionally write them as JSON."""
        records = self.iter_pricing_runs()
        if out_path is not None:
            target = Path(out_path)
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_text(_json_dumps(records) + "\n", encoding="utf-8")
        return records

    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.path)
        conn.row_factory = sqlite3.Row
        return conn


def record_pricing_run(ledger_path: str | Path, summary: Mapping[str, Any]) -> dict[str, Any]:
    """Persist one summary to *ledger_path* using :class:`SQLiteRunLedger`."""
    return SQLiteRunLedger(ledger_path).record_pricing_run(summary)


def export_ledger_json(
    ledger_path: str | Path,
    out_path: str | Path | None = None,
) -> list[dict[str, Any]]:
    """Export all pricing-run summaries from *ledger_path* as JSON."""
    return SQLiteRunLedger(ledger_path).export_json(out_path)


def _normalize_summary(summary: Mapping[str, Any]) -> dict[str, Any]:
    required = (
        "run_id",
        "product",
        "scenario_id",
        "assumption_artifacts",
        "input_hash",
        "output_metrics",
        "parity_status",
        "created_at",
    )
    missing = [key for key in required if key not in summary]
    if missing:
        raise ValueError(f"Pricing run summary is missing required fields: {missing}.")

    normalized = dict(summary)
    normalized["run_id"] = str(normalized["run_id"])
    normalized["product"] = str(normalized["product"])
    normalized["scenario_id"] = str(normalized["scenario_id"])
    normalized["created_at"] = str(normalized["created_at"])
    normalized["input_hash"] = str(normalized["input_hash"])
    normalized["parity_status"] = str(normalized["parity_status"])
    normalized["assumption_artifacts"] = list(normalized["assumption_artifacts"])
    normalized["output_metrics"] = dict(normalized["output_metrics"])
    if "output_paths" in normalized:
        normalized["output_paths"] = [str(path) for path in normalized["output_paths"]]
    if "metadata" in normalized:
        normalized["metadata"] = dict(normalized["metadata"])
    if "assumption_evidence" in normalized:
        normalized["assumption_evidence"] = dict(normalized["assumption_evidence"])
    return normalized


def _json_dumps(payload: Any) -> str:
    return json.dumps(payload, sort_keys=True, default=str, separators=(",", ":"))


__all__ = [
    "SCHEMA_VERSION",
    "SQLiteRunLedger",
    "default_ledger_path",
    "export_ledger_json",
    "pricing_run_summary",
    "record_pricing_run",
    "stable_input_hash",
]
