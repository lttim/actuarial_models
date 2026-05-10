from __future__ import annotations

import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = REPO_ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))

import agent_team_router as router  # noqa: E402


def _role_ids(plan: router.StaffingPlan) -> set[str]:
    return {role.role_id for role in plan.roles}


def test_calculation_and_excel_staffs_model_builder_validation_and_actuary() -> None:
    plan = router.build_staffing_plan(
        [
            "annuity_model/src/annuity_model/rila_projection.py",
            "annuity_model/src/annuity_model/build_rila_excel_workbook.py",
        ],
        objective="change rila mechanics",
    )

    roles = _role_ids(plan)
    assert plan.multi_agent_required
    assert {"calculation", "excel"} <= set(plan.surfaces)
    assert {
        "orchestrator_integrator",
        "model_engineer",
        "excel_builder_engineer",
        "validation_engineer",
        "actuarial_peer_reviewer",
    } <= roles
    assert "parity" in plan.gate_ids
    assert "deep_smoke" in plan.gate_ids
    assert "mutmut_touched" in plan.gate_ids


def test_ui_state_change_staffs_ux_and_validation() -> None:
    plan = router.build_staffing_plan(
        [
            "annuity_model/src/annuity_model/pricing_ui.py",
            "annuity_model/src/annuity_model/pricing_run_form_state.py",
        ]
    )

    roles = _role_ids(plan)
    assert "ui" in plan.surfaces
    assert "ux_reviewer_builder" in roles
    assert "validation_engineer" in roles
    assert "ui_apptest" in plan.gate_ids


def test_docs_only_change_staffs_docs_steward_without_forcing_full_team() -> None:
    plan = router.build_staffing_plan(["annuity_model/docs/index.md"])

    roles = _role_ids(plan)
    assert plan.surfaces == ("docs_governance",)
    assert not plan.multi_agent_required
    assert roles == {"orchestrator_integrator", "docs_governance_steward"}
    assert "docs_inventory" in plan.gate_ids


def test_security_and_packaging_create_dynamic_specialists() -> None:
    plan = router.build_staffing_plan(
        [
            "annuity_model/pyproject.toml",
            "annuity_model/requirements.txt",
            ".github/workflows/security.yml",
        ]
    )

    roles = {role.role_id: role for role in plan.roles}
    assert "packaging_engineer" in roles
    assert "security_reviewer" in roles
    assert roles["packaging_engineer"].dynamic
    assert roles["security_reviewer"].dynamic
    assert "security" in plan.gate_ids


def test_release_waiver_evidence_routes_assumption_and_release_review() -> None:
    plan = router.build_staffing_plan(["annuity_model/.release/assumption_waiver.md"])

    roles = _role_ids(plan)
    assert {"assumptions", "release"} <= set(plan.surfaces)
    assert "data_governance_reviewer" in roles
    assert "release_manager" in roles
    assert "assumption_guardrail" in plan.gate_ids


def test_agent_control_plane_scripts_route_to_validation() -> None:
    plan = router.build_staffing_plan(
        [
            "annuity_model/scripts/agent_preflight.py",
            "annuity_model/scripts/agent_team_router.py",
        ]
    )

    assert "validation" in plan.surfaces
    assert "validation_engineer" in _role_ids(plan)
    assert "full_pytest" in plan.gate_ids


def test_codeowners_routes_to_governance_and_release_control_plane() -> None:
    plan = router.build_staffing_plan([".github/CODEOWNERS"])

    assert {"docs_governance", "release"} <= set(plan.surfaces)
    roles = _role_ids(plan)
    assert "docs_governance_steward" in roles
    assert "release_manager" in roles
    assert "docs_inventory" in plan.gate_ids


def test_broad_mixed_change_adds_migration_planner_and_soft_cap_signal() -> None:
    plan = router.build_staffing_plan(
        [
            "annuity_model/pyproject.toml",
            "annuity_model/src/annuity_model/__init__.py",
            "annuity_model/src/annuity_model/pricing_ui.py",
            "annuity_model/src/annuity_model/rila_projection.py",
            "annuity_model/src/annuity_model/build_rila_excel_workbook.py",
            "annuity_model/docs/AI_AGENT_PREFLIGHT.md",
            "annuity_model/src/annuity_model/data_registry.py",
            "annuity_model/tests/test_rila_projection.py",
        ]
    )

    roles = _role_ids(plan)
    assert "migration_planner" in roles
    assert plan.soft_cap_exceeded
    assert plan.recommended_concurrency == router.SOFT_CONCURRENCY_CAP
    assert any("soft concurrency cap exceeded" in item for item in plan.staffing_rationale)


def test_every_dynamic_role_has_declared_role_contract_fields() -> None:
    plan = router.build_staffing_plan(
        [
            "annuity_model/requirements.txt",
            "annuity_model/src/annuity_model/data/assumptions/assumption_approvals.json",
            "annuity_model/tests/test_perf_baselines.py",
            ".github/branch-protection.json",
        ]
    )

    dynamic_roles = [role for role in plan.roles if role.dynamic]
    assert dynamic_roles
    for role in dynamic_roles:
        assert role.purpose
        assert role.authority
        assert role.expected_artifact
        assert isinstance(role.acceptance_checks, tuple)
        assert role.acceptance_checks
        assert role.rationale
