from __future__ import annotations

import json
import sqlite3

import pytest

from annuity_model.run_ledger import (
    SCHEMA_VERSION,
    SQLiteRunLedger,
    default_ledger_path,
    export_ledger_json,
    pricing_run_summary,
    record_pricing_run,
    stable_input_hash,
)


def _summary(run_id: str = "pricing-0001", *, product: str = "spia") -> dict[str, object]:
    return pricing_run_summary(
        run_id=run_id,
        product=product,
        scenario_id="base",
        assumption_artifacts=[{"artifact_name": "treasury_zero_curve", "hash": "abc"}],
        input_payload={"product": product, "issue_age": 65, "spread": 0.0},
        output_metrics={"single_premium": 100.0},
        parity_status="prepared",
        created_at_utc="2026-05-03T00:00:00Z",
        output_paths=["/tmp/demo.xlsx"],
        validation_status="passed",
        waiver_status="not_required",
        assumption_evidence={"waiver_status": "not_required", "artifact_count": 1},
        model_version="demo-v1",
        git_commit="abcdef0",
        metadata={"source": "unit-test"},
    )


def test_pricing_run_summary_keeps_existing_shape_and_optional_audit_fields() -> None:
    payload = {"product": "spia", "issue_age": 65, "spread": 0.0}
    summary = pricing_run_summary(
        run_id="pricing-0001",
        product="spia",
        scenario_id="base",
        assumption_artifacts=[],
        input_payload=payload,
        output_metrics={"single_premium": 100.0},
        parity_status="prepared",
        created_at_utc="2026-05-03T00:00:00Z",
    )
    assert summary["input_hash"] == stable_input_hash(payload)
    assert "output_paths" not in summary

    enriched = _summary()
    assert enriched["output_paths"] == ["/tmp/demo.xlsx"]
    assert enriched["validation_status"] == "passed"
    assert enriched["assumption_evidence"]["artifact_count"] == 1
    assert enriched["metadata"] == {"source": "unit-test"}


def test_sqlite_run_ledger_persists_and_exports_json(tmp_path) -> None:
    ledger_path = tmp_path / "ledger.sqlite3"
    export_path = tmp_path / "ledger.json"
    ledger = SQLiteRunLedger(ledger_path)

    stored = ledger.record_pricing_run(_summary())
    assert stored["run_id"] == "pricing-0001"
    assert ledger.get_pricing_run("pricing-0001") == stored
    assert ledger.get_pricing_run("missing") is None

    records = ledger.export_json(export_path)
    assert records == [stored]
    assert json.loads(export_path.read_text(encoding="utf-8")) == [stored]

    with sqlite3.connect(ledger_path) as conn:
        row = conn.execute(
            "SELECT schema_version, product, validation_status FROM pricing_runs WHERE run_id = ?",
            ("pricing-0001",),
        ).fetchone()
    assert row == (SCHEMA_VERSION, "spia", "passed")


def test_sqlite_run_ledger_filters_and_replaces_run_ids(tmp_path) -> None:
    ledger = SQLiteRunLedger(tmp_path / "ledger.sqlite3")
    ledger.record_pricing_run(_summary("pricing-0001", product="spia"))
    ledger.record_pricing_run(_summary("pricing-0002", product="term_life"))

    replacement = _summary("pricing-0001", product="spia")
    replacement["parity_status"] = "validated"
    ledger.record_pricing_run(replacement)

    assert [r["run_id"] for r in ledger.iter_pricing_runs()] == ["pricing-0001", "pricing-0002"]
    assert [r["run_id"] for r in ledger.iter_pricing_runs(product="spia")] == ["pricing-0001"]
    assert ledger.get_pricing_run("pricing-0001")["parity_status"] == "validated"


def test_top_level_helpers_and_default_path(tmp_path) -> None:
    ledger_path = tmp_path / "nested" / "ledger.sqlite3"
    stored = record_pricing_run(ledger_path, _summary())
    assert stored["run_id"] == "pricing-0001"
    assert export_ledger_json(ledger_path) == [stored]
    assert default_ledger_path(tmp_path) == tmp_path / "outputs" / "run_ledger.sqlite3"


def test_sqlite_run_ledger_rejects_incomplete_summaries(tmp_path) -> None:
    ledger = SQLiteRunLedger(tmp_path / "ledger.sqlite3")
    with pytest.raises(ValueError, match="missing required fields"):
        ledger.record_pricing_run({"run_id": "pricing-0001"})
