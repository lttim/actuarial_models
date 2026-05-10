"""Autonomous AI-team staffing router.

This module is intentionally deterministic: given a list of changed paths and
an optional task objective, it returns the agent roles that should participate,
their authority, their owned path scopes, and the validation gates the
orchestrator should run before claiming completion.

It does not spawn agents itself. Spawning is host-specific (Codex, Cursor,
Claude Code, etc.). The router produces the auditable contract that the host
orchestrator follows and records in a Team Run Packet.
"""

from __future__ import annotations

import argparse
import fnmatch
import json
from collections.abc import Iterable
from dataclasses import asdict, dataclass

CORE_ROLE_IDS: tuple[str, ...] = (
    "orchestrator_integrator",
    "model_engineer",
    "excel_builder_engineer",
    "validation_engineer",
    "actuarial_peer_reviewer",
    "ux_reviewer_builder",
    "docs_governance_steward",
)

SOFT_CONCURRENCY_CAP = 5


SURFACE_GLOBS: dict[str, tuple[str, ...]] = {
    "calculation": (
        "annuity_model/src/annuity_model/*_projection.py",
        "annuity_model/src/annuity_model/account_value.py",
        "annuity_model/src/annuity_model/crediting.py",
        "annuity_model/src/annuity_model/lapse.py",
        "annuity_model/src/annuity_model/dynamic_lapse.py",
        "annuity_model/src/annuity_model/mortality_2017_cso.py",
        "annuity_model/src/annuity_model/parity_constants.py",
        "annuity_model/src/annuity_model/actuarial_benchmarks.py",
        "annuity_model/src/annuity_model/product_registry.py",
        "annuity_model/src/annuity_model/product_excel.py",
        "annuity_model/src/annuity_model/liability_dispatch.py",
        "annuity_model/src/annuity_model/liability_layouts.py",
        "annuity_model/src/annuity_model/products/*/engine.py",
        "annuity_model/src/annuity_model/products/*/schema.py",
        "annuity_model/src/annuity_model/products/*/excel.py",
    ),
    "excel": (
        "annuity_model/src/annuity_model/build_*_excel_workbook.py",
        "annuity_model/src/annuity_model/build_portfolio_excel_workbook.py",
        "annuity_model/src/annuity_model/alm_excel_ladder.py",
        "annuity_model/src/annuity_model/excel_builder_helpers.py",
        "annuity_model/src/annuity_model/excel_workbook_validator.py",
        "annuity_model/src/annuity_model/product_excel.py",
        "annuity_model/src/annuity_model/liability_layouts.py",
    ),
    "portfolio": (
        "annuity_model/src/annuity_model/portfolio*.py",
        "annuity_model/src/annuity_model/liability_aggregation.py",
        "annuity_model/src/annuity_model/inforce_*.py",
        "annuity_model/src/annuity_model/products/*/inforce.py",
        "annuity_model/tests/parity/portfolio/**",
        "annuity_model/tests/integration/test_portfolio_cli.py",
    ),
    "ui": (
        "streamlit_app.py",
        "annuity_model/src/annuity_model/pricing_ui.py",
        "annuity_model/src/annuity_model/pricing_run_form_state.py",
        "annuity_model/src/annuity_model/ui/**",
        "annuity_model/src/annuity_model/products/*/ui.py",
        "annuity_model/tests/ui/**",
    ),
    "docs_governance": (
        "AGENTS.md",
        "CONTRIBUTING.md",
        "DOCUMENTATION_MAP.md",
        "PROJECT_DEVELOPMENT_GUIDE.md",
        "README.md",
        ".github/CODEOWNERS",
        ".cursor/**",
        "annuity_model/AGENTS.md",
        "annuity_model/README.md",
        "annuity_model/docs/**",
        "actuarial_parity_kit/**",
    ),
    "validation": (
        "annuity_model/tests/**",
        "annuity_model/scripts/agent_*.py",
        "annuity_model/scripts/check_*.py",
        "annuity_model/scripts/deep_smoke.py",
        "annuity_model/scripts/render_*.py",
        "annuity_model/scripts/mutmut_pr_gate.py",
        "annuity_model/mutmut_thresholds.toml",
        ".pre-commit-config.yaml",
        ".github/workflows/**",
    ),
    "assumptions": (
        "annuity_model/.release/assumption_waiver.*",
        "annuity_model/.release/assumptions/**",
        "annuity_model/src/annuity_model/data/**",
        "annuity_model/src/annuity_model/data_registry.py",
        "annuity_model/src/annuity_model/assumption_provenance.py",
        "annuity_model/src/annuity_model/scenario_catalog.py",
        "annuity_model/src/annuity_model/pricing_scenario_materialize.py",
        "annuity_model/docs/assumption_governance.md",
        "annuity_model/docs/release_assumption_waiver.md",
    ),
    "packaging": (
        "annuity_model/pyproject.toml",
        "annuity_model/src/annuity_model/__init__.py",
        "annuity_model/src/annuity_model/launchers.py",
        "annuity_model/pytest.ini",
        "annuity_model/requirements*.txt",
        "requirements.txt",
        "Dockerfile",
        ".devcontainer/**",
        "annuity_model/run_*.sh",
        "annuity_model/run_*.bat",
        "annuity_model/run_*.command",
    ),
    "security": (
        "annuity_model/requirements*.txt",
        "requirements.txt",
        "Dockerfile",
        ".github/workflows/security.yml",
        ".github/dependabot.yml",
    ),
    "release": (
        "annuity_model/.release/**",
        "annuity_model/docs/CHANGELOG.md",
        "annuity_model/docs/model_change_log.md",
        "annuity_model/docs/runbooks/release.md",
        ".github/branch-protection*.json",
        ".github/CODEOWNERS",
        ".github/pull_request_template.md",
    ),
    "performance": (
        "annuity_model/tests/test_perf_baselines.py",
        "annuity_model/tests/benchmarks/**",
        "annuity_model/.benchmarks/**",
        "annuity_model/scripts/deep_assessment.py",
    ),
}


@dataclass(frozen=True, slots=True)
class RoleContract:
    role_id: str
    display_name: str
    purpose: str
    authority: str
    read_only: bool
    owned_paths: tuple[str, ...]
    dependencies: tuple[str, ...]
    expected_artifact: str
    acceptance_checks: tuple[str, ...]
    dynamic: bool
    rationale: str


@dataclass(frozen=True, slots=True)
class StaffingPlan:
    objective: str
    changed_files: tuple[str, ...]
    surfaces: tuple[str, ...]
    roles: tuple[RoleContract, ...]
    multi_agent_required: bool
    recommended_concurrency: int
    soft_cap_exceeded: bool
    staffing_rationale: tuple[str, ...]
    gate_ids: tuple[str, ...]

    def to_dict(self) -> dict[str, object]:
        payload = asdict(self)
        payload["roles"] = [asdict(role) for role in self.roles]
        return payload


@dataclass(frozen=True, slots=True)
class GateSpec:
    gate_id: str
    label: str
    command: tuple[str, ...]
    cwd: str
    env: dict[str, str]
    rationale: str

    def to_dict(self) -> dict[str, object]:
        return asdict(self)


GATES: dict[str, GateSpec] = {
    "docs_inventory": GateSpec(
        "docs_inventory",
        "Documentation inventory",
        ("python", "scripts/check_documentation_map.py"),
        "annuity_model",
        {},
        "Tracked markdown files and DOCUMENTATION_MAP.md must stay aligned.",
    ),
    "parity_contract": GateSpec(
        "parity_contract",
        "Rendered parity contracts",
        ("python", "scripts/render_parity_contract.py", "--check"),
        "annuity_model",
        {},
        "Tolerance tables render from parity_constants.py.",
    ),
    "actuarial_benchmarks": GateSpec(
        "actuarial_benchmarks",
        "Rendered actuarial benchmarks",
        ("python", "scripts/render_actuarial_benchmarks.py", "--check"),
        "annuity_model",
        {},
        "Benchmark documentation must match actuarial_benchmarks.py.",
    ),
    "parity": GateSpec(
        "parity",
        "Python/Excel parity",
        ("python", "-m", "pytest", "tests/parity", "-q"),
        "annuity_model",
        {},
        "Calculation and workbook changes must preserve parity.",
    ),
    "full_pytest": GateSpec(
        "full_pytest",
        "Full pytest suite",
        ("python", "-m", "pytest", "-q"),
        "annuity_model",
        {},
        "Behavior changes require the full regression suite.",
    ),
    "deep_smoke": GateSpec(
        "deep_smoke",
        "Deep smoke workbook validation",
        ("python", "scripts/deep_smoke.py"),
        "annuity_model",
        {},
        "Every implemented product workbook must build and validate.",
    ),
    "ui_apptest": GateSpec(
        "ui_apptest",
        "Streamlit AppTest suite",
        ("python", "-m", "pytest", "tests/ui", "-q"),
        "annuity_model",
        {},
        "UI changes must render and run workflows under AppTest.",
    ),
    "portfolio_acceptance_subset": GateSpec(
        "portfolio_acceptance_subset",
        "Portfolio acceptance subset",
        ("python", "-m", "pytest", "tests/parity/portfolio", "tests/integration", "-q"),
        "annuity_model",
        {"ANNUITY_MODEL_PORTFOLIO_V1": "1"},
        "Portfolio changes require portfolio parity and CLI integration evidence.",
    ),
    "assumption_guardrail": GateSpec(
        "assumption_guardrail",
        "Assumption release guardrail",
        ("python", "scripts/check_assumption_release_guardrails.py"),
        "annuity_model",
        {},
        "Placeholder/synthetic assumptions must be blocked or explicitly waived.",
    ),
    "pre_commit": GateSpec(
        "pre_commit",
        "Pre-commit hooks",
        ("pre-commit", "run", "--all-files"),
        ".",
        {},
        "Lint, format, type, and local policy hooks must pass.",
    ),
    "security": GateSpec(
        "security",
        "Security scan",
        ("just", "security"),
        ".",
        {},
        "Dependency and static security checks for packaging/security changes.",
    ),
    "mutmut_touched": GateSpec(
        "mutmut_touched",
        "Touched-file mutation gate",
        ("python", "scripts/mutmut_pr_gate.py", "--touched-files", "{changed_files}"),
        "annuity_model",
        {},
        "Parity-critical touched files must stay within mutation survivor thresholds.",
    ),
}


def _match_any(path: str, patterns: Iterable[str]) -> bool:
    return any(fnmatch.fnmatch(path, pattern) for pattern in patterns)


def classify_changed_files(changed_files: Iterable[str]) -> tuple[str, ...]:
    """Return sorted surface ids touched by *changed_files*."""
    surfaces: set[str] = set()
    for path in changed_files:
        for surface, globs in SURFACE_GLOBS.items():
            if _match_any(path, globs):
                surfaces.add(surface)
    return tuple(sorted(surfaces))


def _role_catalog() -> dict[str, RoleContract]:
    return {
        "orchestrator_integrator": RoleContract(
            role_id="orchestrator_integrator",
            display_name="Orchestrator / Integrator",
            purpose="Decompose the task, staff agents, integrate outputs, run final gates, and declare completion.",
            authority="integration-owner",
            read_only=False,
            owned_paths=("*",),
            dependencies=(),
            expected_artifact="Team Run Packet with staffing, decisions, gates, and final signoff.",
            acceptance_checks=(
                "all selected gates pass",
                "no unresolved blocking reviewer findings",
            ),
            dynamic=False,
            rationale="Every autonomous run needs one accountable integration owner.",
        ),
        "model_engineer": RoleContract(
            role_id="model_engineer",
            display_name="Model Engineer",
            purpose="Implement scoped Python model, product-definition, liability, or scenario changes.",
            authority="scoped-write",
            read_only=False,
            owned_paths=(
                "annuity_model/src/annuity_model/*_projection.py",
                "annuity_model/src/annuity_model/products/*/(engine|schema).py",
                "annuity_model/src/annuity_model/product_registry.py",
                "annuity_model/src/annuity_model/liability_*.py",
            ),
            dependencies=("actuarial_peer_reviewer", "validation_engineer"),
            expected_artifact="Implementation summary plus model/parity risk notes.",
            acceptance_checks=(
                "parity",
                "full_pytest",
                "actuarial peer review when calculation-facing",
            ),
            dynamic=False,
            rationale="Calculation-facing changes require a focused model implementer.",
        ),
        "excel_builder_engineer": RoleContract(
            role_id="excel_builder_engineer",
            display_name="Excel Builder Engineer",
            purpose="Implement scoped workbook-builder, formula, layout, and validator-facing changes.",
            authority="scoped-write",
            read_only=False,
            owned_paths=(
                "annuity_model/src/annuity_model/build_*_excel_workbook.py",
                "annuity_model/src/annuity_model/alm_excel_ladder.py",
                "annuity_model/src/annuity_model/excel_*",
                "annuity_model/src/annuity_model/liability_layouts.py",
            ),
            dependencies=("validation_engineer",),
            expected_artifact="Workbook-change summary with validator and ModelCheck implications.",
            acceptance_checks=("parity", "deep_smoke", "tests/test_excel_export_validation.py"),
            dynamic=False,
            rationale="Excel is an independent audit surface and needs separate implementation attention.",
        ),
        "validation_engineer": RoleContract(
            role_id="validation_engineer",
            display_name="Validation Engineer",
            purpose="Design and run tests, parity checks, smoke checks, mutation/security evidence, and regression coverage.",
            authority="scoped-write",
            read_only=False,
            owned_paths=(
                "annuity_model/tests/**",
                "annuity_model/scripts/check_*.py",
                "annuity_model/scripts/agent_*.py",
                "annuity_model/scripts/*_gate.py",
            ),
            dependencies=(),
            expected_artifact="Validation matrix with commands, exit codes, and residual risk.",
            acceptance_checks=("selected gates produce exit code 0", "new behavior has tests"),
            dynamic=False,
            rationale="Autonomous work must include an agent whose main job is proving the work.",
        ),
        "actuarial_peer_reviewer": RoleContract(
            role_id="actuarial_peer_reviewer",
            display_name="Actuarial Peer Reviewer",
            purpose="Readonly actuarial review of methodology, assumptions, outputs, and product fitness.",
            authority="read-only-review",
            read_only=True,
            owned_paths=(),
            dependencies=("validation_engineer",),
            expected_artifact="Actuarial verdict or SME review reference.",
            acceptance_checks=("no BLOCK findings", "human-judgment findings escalated"),
            dynamic=False,
            rationale="Tests cannot detect actuarially coherent but methodologically weak changes.",
        ),
        "ux_reviewer_builder": RoleContract(
            role_id="ux_reviewer_builder",
            display_name="UX Reviewer / Builder",
            purpose="Implement scoped Streamlit UI changes and assess demo usability.",
            authority="scoped-write",
            read_only=False,
            owned_paths=(
                "streamlit_app.py",
                "annuity_model/src/annuity_model/pricing_ui.py",
                "annuity_model/src/annuity_model/ui/**",
                "annuity_model/src/annuity_model/products/*/ui.py",
                "annuity_model/tests/ui/**",
            ),
            dependencies=("validation_engineer",),
            expected_artifact="UX change summary with AppTest evidence and demo-flow notes.",
            acceptance_checks=("ui_apptest", "full_pytest"),
            dynamic=False,
            rationale="Professional-demo quality depends on a deliberate user experience surface.",
        ),
        "docs_governance_steward": RoleContract(
            role_id="docs_governance_steward",
            display_name="Docs / Governance Steward",
            purpose="Keep AI instructions, docs inventory, governance runbooks, and release evidence aligned.",
            authority="scoped-write",
            read_only=False,
            owned_paths=(
                "AGENTS.md",
                "PROJECT_DEVELOPMENT_GUIDE.md",
                "DOCUMENTATION_MAP.md",
                "annuity_model/AGENTS.md",
                "annuity_model/docs/**",
            ),
            dependencies=(),
            expected_artifact="Documentation and governance alignment summary.",
            acceptance_checks=(
                "docs_inventory",
                "parity_contract when tolerance docs are in scope",
            ),
            dynamic=False,
            rationale="Future AI agents rely on current docs as operational controls.",
        ),
        "security_reviewer": RoleContract(
            role_id="security_reviewer",
            display_name="Security Reviewer",
            purpose="Readonly or scoped review of dependency, workflow, container, and static security surfaces.",
            authority="read-only-review",
            read_only=True,
            owned_paths=(),
            dependencies=("validation_engineer",),
            expected_artifact="Security scan summary and dependency risk notes.",
            acceptance_checks=("security",),
            dynamic=True,
            rationale="Security expertise is task-specific and only needed for dependency/container/control-plane changes.",
        ),
        "packaging_engineer": RoleContract(
            role_id="packaging_engineer",
            display_name="Packaging Engineer",
            purpose="Implement scoped package layout, launcher, dependency, and import-surface changes.",
            authority="scoped-write",
            read_only=False,
            owned_paths=(
                "annuity_model/pyproject.toml",
                "annuity_model/src/annuity_model/__init__.py",
                "annuity_model/src/annuity_model/launchers.py",
                "annuity_model/pytest.ini",
                "annuity_model/requirements*.txt",
                "requirements.txt",
                "Dockerfile",
                "annuity_model/run_*",
            ),
            dependencies=("validation_engineer",),
            expected_artifact="Packaging/import compatibility summary.",
            acceptance_checks=("full_pytest", "launcher invariants", "import smoke tests"),
            dynamic=True,
            rationale="Packaging work needs import and launcher discipline distinct from model work.",
        ),
        "data_governance_reviewer": RoleContract(
            role_id="data_governance_reviewer",
            display_name="Data Governance Reviewer",
            purpose="Review assumption artifacts, approval metadata, placeholder status, and waiver requirements.",
            authority="read-only-review",
            read_only=True,
            owned_paths=(),
            dependencies=("docs_governance_steward",),
            expected_artifact="Assumption-governance findings and waiver requirements.",
            acceptance_checks=("assumption_guardrail",),
            dynamic=True,
            rationale="Assumption changes require governance review even when formulas do not change.",
        ),
        "performance_engineer": RoleContract(
            role_id="performance_engineer",
            display_name="Performance Engineer",
            purpose="Assess benchmark impact and runtime budget for expensive model/workbook operations.",
            authority="scoped-write",
            read_only=False,
            owned_paths=(
                "annuity_model/tests/test_perf_baselines.py",
                "annuity_model/tests/benchmarks/**",
            ),
            dependencies=("validation_engineer",),
            expected_artifact="Benchmark delta and performance-risk summary.",
            acceptance_checks=("performance benchmarks pass or deltas are justified",),
            dynamic=True,
            rationale="Performance work is specialized and only needed when runtime-sensitive surfaces change.",
        ),
        "release_manager": RoleContract(
            role_id="release_manager",
            display_name="Release Manager",
            purpose="Coordinate release notes, branch protection, waiver evidence, and final release readiness.",
            authority="scoped-write",
            read_only=False,
            owned_paths=(
                "annuity_model/docs/CHANGELOG.md",
                "annuity_model/docs/model_change_log.md",
                ".github/branch-protection*.json",
                ".github/pull_request_template.md",
            ),
            dependencies=("validation_engineer", "docs_governance_steward"),
            expected_artifact="Release readiness checklist and residual-risk summary.",
            acceptance_checks=(
                "selected release gates pass",
                "waiver evidence present when required",
            ),
            dynamic=True,
            rationale="Release work has separate evidence and branch-protection concerns.",
        ),
        "migration_planner": RoleContract(
            role_id="migration_planner",
            display_name="Migration Planner",
            purpose="Plan and review broad package/UI/product-definition migrations for sequencing and compatibility.",
            authority="read-only-review",
            read_only=True,
            owned_paths=(),
            dependencies=("orchestrator_integrator",),
            expected_artifact="Migration sequence, compatibility risks, and rollback notes.",
            acceptance_checks=("migration has staged tests and compatibility gates",),
            dynamic=True,
            rationale="Large refactors need a planning reviewer separate from implementers.",
        ),
    }


def _role_ids_for_surfaces(surfaces: set[str], changed_files: tuple[str, ...]) -> list[str]:
    role_ids = ["orchestrator_integrator"]

    if surfaces & {"calculation", "portfolio"}:
        role_ids.append("model_engineer")
    if "excel" in surfaces:
        role_ids.append("excel_builder_engineer")
    if surfaces & {"calculation", "excel", "portfolio", "assumptions"}:
        role_ids.append("actuarial_peer_reviewer")
    if "ui" in surfaces:
        role_ids.append("ux_reviewer_builder")
    if surfaces & {"docs_governance", "release"}:
        role_ids.append("docs_governance_steward")
    if surfaces & {"validation", "calculation", "excel", "portfolio", "ui", "packaging"}:
        role_ids.append("validation_engineer")
    if "assumptions" in surfaces:
        role_ids.append("data_governance_reviewer")
    if "security" in surfaces:
        role_ids.append("security_reviewer")
    if "packaging" in surfaces:
        role_ids.append("packaging_engineer")
    if "performance" in surfaces:
        role_ids.append("performance_engineer")
    if "release" in surfaces:
        role_ids.append("release_manager")

    broad_refactor = len(surfaces) >= 4 or len(changed_files) >= 8
    if (
        broad_refactor
        or ({"packaging", "ui"} <= surfaces)
        or ({"packaging", "calculation"} <= surfaces)
    ):
        role_ids.append("migration_planner")

    deduped: list[str] = []
    for role_id in role_ids:
        if role_id not in deduped:
            deduped.append(role_id)
    return deduped


def _gate_ids_for_surfaces(surfaces: set[str], multi_agent_required: bool) -> tuple[str, ...]:
    gates: list[str] = []
    if surfaces & {"docs_governance", "release"}:
        gates.extend(["docs_inventory", "parity_contract"])
    if surfaces & {"calculation", "excel", "portfolio"}:
        gates.extend(
            [
                "parity",
                "full_pytest",
                "deep_smoke",
                "parity_contract",
                "actuarial_benchmarks",
                "mutmut_touched",
            ]
        )
    if "ui" in surfaces:
        gates.extend(["ui_apptest", "full_pytest"])
    if "portfolio" in surfaces:
        gates.extend(["portfolio_acceptance_subset"])
    if "assumptions" in surfaces:
        gates.extend(["assumption_guardrail", "docs_inventory"])
    if "packaging" in surfaces:
        gates.extend(["full_pytest", "deep_smoke"])
    if "security" in surfaces:
        gates.extend(["security"])
    if "validation" in surfaces:
        gates.extend(["full_pytest"])
    if multi_agent_required and not gates:
        gates.extend(["full_pytest"])

    deduped: list[str] = []
    for gate in gates:
        if gate not in deduped:
            deduped.append(gate)
    return tuple(deduped)


def build_staffing_plan(
    changed_files: Iterable[str],
    *,
    objective: str = "",
) -> StaffingPlan:
    changed = tuple(sorted({p for p in changed_files if p}))
    surfaces = set(classify_changed_files(changed))
    high_risk = bool(surfaces & {"calculation", "excel", "portfolio", "assumptions", "packaging"})
    multi_subsystem = len(surfaces) >= 2
    broad_change = len(changed) >= 6
    multi_agent_required = high_risk or multi_subsystem or broad_change

    catalog = _role_catalog()
    role_ids = _role_ids_for_surfaces(surfaces, changed)
    if multi_agent_required and "validation_engineer" not in role_ids:
        role_ids.append("validation_engineer")
    roles = tuple(catalog[role_id] for role_id in role_ids)

    active_roles = [role for role in roles if role.role_id != "orchestrator_integrator"]
    recommended_concurrency = min(SOFT_CONCURRENCY_CAP, max(1, len(active_roles)))
    soft_cap_exceeded = len(active_roles) > SOFT_CONCURRENCY_CAP

    rationale: list[str] = []
    if high_risk:
        rationale.append("high-risk actuarial/platform surface detected")
    if multi_subsystem:
        rationale.append(f"multiple surfaces detected: {', '.join(sorted(surfaces))}")
    if broad_change:
        rationale.append(f"broad change set detected: {len(changed)} changed files")
    if not rationale:
        rationale.append(
            "single-surface low-risk change; orchestrator may handle without a full team"
        )
    if soft_cap_exceeded:
        rationale.append(
            f"soft concurrency cap exceeded: {len(active_roles)} non-orchestrator roles selected; "
            f"run at most {SOFT_CONCURRENCY_CAP} in parallel unless write scopes are independent"
        )

    gate_ids = _gate_ids_for_surfaces(surfaces, multi_agent_required)
    return StaffingPlan(
        objective=objective,
        changed_files=changed,
        surfaces=tuple(sorted(surfaces)),
        roles=roles,
        multi_agent_required=multi_agent_required,
        recommended_concurrency=recommended_concurrency,
        soft_cap_exceeded=soft_cap_exceeded,
        staffing_rationale=tuple(rationale),
        gate_ids=gate_ids,
    )


def gate_specs_for_plan(plan: StaffingPlan) -> tuple[GateSpec, ...]:
    return tuple(GATES[gate_id] for gate_id in plan.gate_ids)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Route changed files to an autonomous AI team.")
    parser.add_argument(
        "--objective", default="", help="Short task objective for the staffing plan."
    )
    parser.add_argument(
        "--changed-files", nargs="*", default=(), help="Repo-relative changed paths."
    )
    parser.add_argument(
        "--json", action="store_true", help="Print JSON instead of a markdown summary."
    )
    args = parser.parse_args(argv)

    plan = build_staffing_plan(args.changed_files, objective=args.objective)
    if args.json:
        print(json.dumps(plan.to_dict(), indent=2, sort_keys=True))
    else:
        print(f"Objective: {plan.objective or '(not provided)'}")
        print(f"Surfaces: {', '.join(plan.surfaces) or '(none detected)'}")
        print(f"Multi-agent required: {plan.multi_agent_required}")
        print(f"Recommended concurrency: {plan.recommended_concurrency}")
        print("Roles:")
        for role in plan.roles:
            dyn = "dynamic" if role.dynamic else "core"
            ro = "read-only" if role.read_only else role.authority
            print(f"  - {role.display_name} [{dyn}; {ro}]")
        print(f"Gates: {', '.join(plan.gate_ids) or '(none selected)'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
