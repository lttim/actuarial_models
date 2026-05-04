# Documentation Map

Complete catalog of tracked documentation in this repository, with one-line purpose per file.

## Workspace-level governance and onboarding

- `README.md` — top-level workspace overview, bootstrap commands, and daily usage.
- `PROJECT_DEVELOPMENT_GUIDE.md` — governance/control framework and cross-platform dev workflow map.
- `AGENTS.md` — workspace-level invariants and cross-platform agent operating rules.
- `CONTRIBUTING.md` — contributor checklist, hard rules, and pre-PR expectations.
- `MACOS_HANDOFF.md` — macOS bootstrap and parity-oriented handoff notes.
- `.github/pull_request_template.md` — PR checklist for products, gates, and parity controls.
- `DOCUMENTATION_MAP.md` — this file; complete documentation inventory.

## Cursor/session workflow docs (tracked helpers)

- `.cursor/handoffs/README.md` — how `!handoff` / `!recall` files work.
- `.cursor/actuary-reviews/README.md` — Actuary SME verdict artifact format and lifecycle.

## Reusable parity-kit templates

- `actuarial_parity_kit/README.md` — parity-kit purpose and setup flow for new repos.
- `actuarial_parity_kit/AGENTS_template.md` — template AGENTS rules for new actuarial products.
- `actuarial_parity_kit/docs/model_parity_contract_template.md` — template parity-contract structure.
- `actuarial_parity_kit/docs/parity_test_checklist_template.md` — template release/parity checklist.

## Annuity-model product-level onboarding and state

- `annuity_model/README.md` — product package architecture, module map, and developer commands.
- `annuity_model/AGENTS.md` — canonical four gates plus product-critical rules.
- `annuity_model/state.md` — compact local snapshot/handoff state for the product folder.
- `annuity_model/ui/MIGRATION.md` — planned decomposition path for `pricing_ui.py`.

## Annuity-model AI-review/skill artifacts (tracked)

- `annuity_model/.cursor/skills/actuary-sme/SKILL.md` — Actuary SME rubric and verdict protocol.
- `annuity_model/.cursor/actuary-reviews/iter-1-2026-04-19-product-portfolio.md` — historical sample SME verdict artifact.

## Data artifact documentation

- `annuity_model/data/mortality/cso_2017_ult/README.md` — synthetic CSO placeholder description and production overlay guidance.

## Annuity-model docs hub

- `annuity_model/docs/index.md` — entry index for contracts, runbooks, governance docs, and roadmaps.
- `annuity_model/docs/AI_AGENT_PREFLIGHT.md` — AI decision tree and source-of-truth map before edits.
- `annuity_model/docs/AI_AGENT_TEAM_PROTOCOL.md` — autonomous multi-agent staffing, role authority, and Team Run Packet protocol.

## Contracts, specs, and release gates

- `annuity_model/docs/model_parity_contract.md` — SPIA/ALM parity contract and tolerance table.
- `annuity_model/docs/rila_parity_contract.md` — RILA-specific parity contract addendum.
- `annuity_model/docs/portfolio_parity_contract.md` — portfolio aggregation parity addendum.
- `annuity_model/docs/rila_product_spec.md` — RILA mechanics-production product behavior/spec definition.
- `annuity_model/docs/iul_product_spec.md` — IUL mechanics-production product behavior/spec definition.
- `annuity_model/docs/portfolio_runner_spec.md` — multi-policy portfolio runner behavior/spec definition.
- `annuity_model/docs/parity_test_checklist.md` — merge/release parity checklist.
- `annuity_model/docs/release_assumption_waiver.md` — waiver template when placeholder assumptions are used.
- `annuity_model/.release/assumption_waiver.md` — active release waiver evidence for placeholder assumptions.

## Governance and controls documentation

- `annuity_model/docs/CHANGELOG.md` — engineering change history for the annuity model package.
- `annuity_model/docs/model_change_log.md` — model-impacting/tolerance-impacting change evidence log.
- `annuity_model/docs/CODEOWNERS_RATIONALE.md` — why code-ownership protections exist by path.
- `annuity_model/docs/model_inventory.md` — model tiering, ownership, intended use, and validation cadence.
- `annuity_model/docs/assumption_governance.md` — assumption metadata/control standard.
- `annuity_model/docs/platform_gap_register.md` — current-vs-target platform maturity register.
- `annuity_model/docs/independent_challenge_activation.md` — second-reviewer activation checklist.

## Product/methodology reference docs

- `annuity_model/docs/glossary.md` — term definitions for actuarial and platform vocabulary.
- `annuity_model/docs/lapse_framework.md` — static lapse v1 framework and limitations.
- `annuity_model/docs/actuarial_benchmarks.md` — benchmark-band rationale with generated constants table.
- `annuity_model/docs/seven_product_rollout_plan.md` — phased rollout blueprint and completion criteria.

## Forward-planning docs

- `annuity_model/docs/actuarial_fidelity_backlog.md` — actuarial capability backlog with evidence expectations.
- `annuity_model/docs/platform_engineering_roadmap.md` — architecture/platform hardening roadmap tracks.
- `annuity_model/docs/implementation_phasing.md` — phased delivery/evidence plan across governance, actuarial, platform work.

## Operational runbooks

- `annuity_model/docs/runbooks/investigate_parity_break.md` — triage workflow for parity failures.
- `annuity_model/docs/runbooks/debug_validator_failure.md` — triage workflow for workbook validator failures.
- `annuity_model/docs/runbooks/regenerate_excel_cache.md` — formula-cache/recalc workflow for Excel workbooks.
- `annuity_model/docs/runbooks/portfolio_run.md` — CLI/UI and acceptance flow for portfolio runs.
- `annuity_model/docs/runbooks/release.md` — release process, tagging, and branch-protection refresh workflow.
- `annuity_model/docs/runbooks/assumption_release_guardrail.md` — release guardrail and waiver usage process.
- `annuity_model/docs/runbooks/launcher_double_click.md` — troubleshooting double-click launcher failures.
