"""Small durable-run summary helpers for demo diagnostics.

This module does not persist records yet; it defines the stable summary shape
the UI and diagnostics export can show today and a file-backed ledger can store
later without changing the public payload.
"""

from __future__ import annotations

import datetime as _dt
import hashlib
import json
from collections.abc import Mapping
from typing import Any


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
) -> dict[str, Any]:
    return {
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


__all__ = ["pricing_run_summary", "stable_input_hash"]
