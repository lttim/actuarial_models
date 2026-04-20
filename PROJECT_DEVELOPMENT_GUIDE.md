# Project Development Guide

This guide explains how the actuarial modeling project is operated end-to-end:

- For **human contributors** who need a practical map of the docs, controls, and release gates.
- For **AI coding agents** (Cursor, Claude Code, Codex, and similar platforms) that must continue work safely with minimal repo-specific assumptions.

Use this file as the navigation hub for development governance. Product math contracts remain in product docs.

## 1) What this project is

`Code_Sandbox` is a multi-surface actuarial platform centered on `annuity_model/`:

- Pricing engines and ALM logic in Python.
- Generated Excel workbooks as independently recalculating auditors.
- Parity, validator, and regression tests that block unsafe changes.
- Operational docs and runbooks for investigations, releases, and roadmap governance.

Main active surfaces:

- `annuity_model/`: production code, tests, docs, runbooks.
- `actuarial_parity_kit/`: reusable governance template for future product repos.
- Root governance docs: `AGENTS.md`, `CONTRIBUTING.md`, this file.

## 2) Documentation map (read in this order)

### Human onboarding

1. `README.md` (workspace overview and bootstrap)
2. `CONTRIBUTING.md` (contribution policy and hard rules)
3. `AGENTS.md` (workspace-wide invariants and cross-platform setup)
4. `annuity_model/README.md` (module map and architecture)
5. `annuity_model/docs/index.md` (runbooks, contracts, governance docs)

### AI onboarding (platform-agnostic)

1. `annuity_model/docs/AI_AGENT_PREFLIGHT.md` (decision tree + source-of-truth map)
2. `annuity_model/AGENTS.md` (canonical completion gates and product-critical rules)
3. Root and product rule files:
   - `.cursor/rules/actuary-sme-protocol.mdc`
   - `.cursor/rules/handoff-recall.mdc`
   - `annuity_model/.cursor/rules/actuarial-parity.mdc`
   - `annuity_model/.cursor/rules/excel-formula-safety.mdc`
4. Relevant runbook under `annuity_model/docs/runbooks/`

If two docs appear inconsistent, resolve in this order:

1. Product/runtime source code contract (`parity_constants.py`, registries, validators)
2. `annuity_model/AGENTS.md` canonical gates section
3. `annuity_model/docs/AI_AGENT_PREFLIGHT.md` routing guidance
4. Secondary docs (`README`, checklists, runbooks)

## 3) Control framework (how quality is enforced)

Development is controlled by layered safeguards. A change is acceptable only when all applicable layers stay green.

1. **Design contracts**
   - Parity contracts: `annuity_model/docs/model_parity_contract.md`, `annuity_model/docs/rila_parity_contract.md`, `annuity_model/docs/portfolio_parity_contract.md`
   - Product specs: `annuity_model/docs/rila_product_spec.md`, `annuity_model/docs/portfolio_runner_spec.md`
2. **Static guards**
   - Formula safety: `annuity_model/excel_workbook_validator.py`
   - Rule constraints in `.cursor/rules/*.mdc`
3. **Regression gates**
   - Parity suite: `annuity_model/tests/parity/`
   - Unit and integration suites: `annuity_model/tests/`
   - End-to-end smoke: `annuity_model/scripts/deep_smoke.py`
4. **Governance coupling**
   - Tolerance source of truth: `annuity_model/parity_constants.py`
   - Required change log updates: `annuity_model/docs/model_change_log.md`
   - Rendered contract consistency: `annuity_model/scripts/render_parity_contract.py --check`
5. **Ownership and review routing**
   - CODEOWNERS path protections with rationale in `annuity_model/docs/CODEOWNERS_RATIONALE.md`
6. **Actuarial judgment loop**
   - Structured Actuary SME review orchestration in `.cursor/rules/actuary-sme-protocol.mdc`
   - SME rubric in `annuity_model/.cursor/skills/actuary-sme/SKILL.md`

## 4) Regression strategy (what "safe change" means)

Regression handling is mandatory, not optional:

- Every numerical bug fix must add a permanent regression test.
- Step-level parity is required where applicable, not just final-value parity.
- Tolerances are never widened as a shortcut to green tests.
- Excel-generating changes require both static validation and runtime recalc confidence paths.

Primary regression references:

- `annuity_model/tests/parity/`
- `annuity_model/docs/parity_test_checklist.md`
- `annuity_model/docs/runbooks/investigate_parity_break.md`
- `annuity_model/docs/runbooks/debug_validator_failure.md`
- `annuity_model/docs/runbooks/runtime_excel_recalc_gate.md`

## 5) Rules, skills, and subagents across AI platforms

This repository uses Cursor-native rule/skill files, but the operating intent is portable.

### Rules (`.mdc`) -> platform-agnostic meaning

- Treat each rule file as a binding policy document (trigger conditions + required actions).
- If your platform does not auto-load `.mdc`, manually read relevant rules before edits.

### Skills (`SKILL.md`) -> platform-agnostic meaning

- Treat skills as specialized standard-operating-procedure modules.
- The Actuary SME skill defines the actuarial review rubric and verdict shape.

### Subagents / delegated reviewers -> platform-agnostic meaning

- Cursor "subagent" maps to any delegated specialist execution model (parallel worker, reviewer agent, tool-running child task).
- The control requirement is invariant: delegated reviews must remain readonly for judgment roles, and parent flow applies fixes with full traceability.

## 6) Canonical completion protocol (all platforms)

Before claiming task completion, run (or provide evidence from CI of) the canonical gates in `annuity_model/AGENTS.md`:

1. `python -m pytest tests/parity -q`
2. `python -m pytest -q`
3. `python scripts/deep_smoke.py`
4. `python scripts/render_parity_contract.py --check`

Plus applicable supersets, e.g. portfolio acceptance (`just portfolio-acceptance`) when portfolio surfaces are touched.

For CALCULATION or TOLERANCE changes, run the Actuary SME review flow and keep verdict artifacts traceable.

## 7) Session continuity and handoffs

- Use `.cursor/rules/handoff-recall.mdc` semantics as the canonical handoff protocol.
- If your AI platform has no native `!handoff` command, generate an equivalent handoff file with:
  - objective
  - current status
  - files touched
  - key commands/results
  - open items
  - exact next step

Continuity quality requirement is platform-independent even when command syntax differs.

## 8) Quick reference by task type

- **Calculation logic change** -> `AI_AGENT_PREFLIGHT.md` CALCULATION branch + parity/investigation runbooks.
- **Tolerance update** -> `AI_AGENT_PREFLIGHT.md` TOLERANCE branch + `model_change_log.md`.
- **Excel builder change** -> `excel-formula-safety.mdc` + validator runbook + parity gates.
- **Portfolio surface change** -> `annuity_model/docs/portfolio_runner_spec.md` + `annuity_model/docs/runbooks/portfolio_run.md`.
- **Docs-only change** -> `just docs-check` from repo root; if root or kit `AGENTS.md` text changes, also run `annuity_model/tests/test_kit_template_parity.py`.
- **Release preparation** -> `annuity_model/docs/runbooks/release.md`.

## 9) Definition of done for maintainable development

A change is done only when:

1. It preserves (or improves) parity and control guarantees.
2. It leaves an audit trail (tests, change logs, verdict artifacts where required).
3. It is understandable by both a human reviewer and a generic AI agent without relying on hidden chat context.
4. It keeps the documented source-of-truth hierarchy intact.
