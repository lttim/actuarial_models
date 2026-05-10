from __future__ import annotations

import numpy as np

from annuity_model.assumption_provenance import (
    approvals_as_dicts,
    provenance_rows_from_pricing_state,
)
from annuity_model.dynamic_lapse import (
    DynamicLapseConfig,
    dynamic_lapse_path,
    persistency_from_monthly_lapse,
)
from annuity_model.experience_study import sample_experience_rows
from annuity_model.run_ledger import pricing_run_summary, stable_input_hash
from annuity_model.scenario_catalog import get_pricing_scenario, list_pricing_scenarios


def test_scenario_catalog_has_replayable_named_stresses() -> None:
    scenarios = list_pricing_scenarios()
    assert len(scenarios) >= 5
    ids = {s.scenario_id for s in scenarios}
    assert {"base", "rates_up_100", "longevity_plus_10", "equity_downturn"} <= ids
    assert get_pricing_scenario("rates_up_100").rate_shift_bps == 100.0
    for scenario in scenarios:
        d = scenario.to_dict()
        assert d["scenario_id"]
        assert d["owner"]
        assert d["intended_use"]


def test_assumption_provenance_surfaces_registry_and_governance_metadata() -> None:
    approvals = approvals_as_dicts()
    assert any(a["artifact_name"] == "treasury_zero_curve" for a in approvals)
    rows = provenance_rows_from_pricing_state(
        pricing_meta={
            "yield_mode": "par_bootstrap",
            "mortality_mode": "rp2014_mp2016",
            "expense_mode": "csv",
            "index_scenario_csv_path": "",
        },
        pricing_run_inputs={},
        pricing_excel_context={},
    )
    names = {r["artifact_name"] for r in rows}
    assert "treasury_zero_curve" in names
    assert "rp2014_male_healthy_annuitant_qx" in names
    assert all("status" in r for r in rows)


def test_run_summary_hash_is_stable_and_changes_with_inputs() -> None:
    payload = {"product": "spia", "issue_age": 65, "spread": 0.0}
    assert stable_input_hash(payload) == stable_input_hash(dict(reversed(payload.items())))
    assert stable_input_hash(payload) != stable_input_hash({**payload, "spread": 0.01})
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
    assert summary["run_id"] == "pricing-0001"
    assert summary["input_hash"] == stable_input_hash(payload)


def test_dynamic_lapse_demo_path_bounds_and_persistency() -> None:
    q = dynamic_lapse_path(
        n_months=240,
        config=DynamicLapseConfig(base_annual_rate=0.04, floor=0.005, cap=0.35),
        rate_shock_bps=150,
        moneyness=0.20,
    )
    assert q.shape == (240,)
    assert np.all(q >= 0.0)
    assert np.all(q < 1.0)
    persistency = persistency_from_monthly_lapse(q)
    assert np.all(np.diff(persistency) <= 1e-12)
    assert 0.0 < float(persistency[-1]) < 1.0


def test_sample_experience_study_contains_review_triggers() -> None:
    rows = sample_experience_rows()
    assert rows
    assert any(r["review_flag"] == "Recommend assumption review" for r in rows)
    assert all(r["claims_oe"] > 0.0 and r["lapse_oe"] > 0.0 for r in rows)
