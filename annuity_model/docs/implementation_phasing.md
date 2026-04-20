# Implementation phasing and evidence plan

This plan sequences roadmap delivery into practical phases with explicit
evidence artifacts.

## Phase 1 (weeks 1-4): governance foundations

### Deliverables

- Baseline gap register publication.
- Model inventory and tiering register.
- Assumption governance standard.
- Placeholder assumption release guardrail script and runbook.
- Second-reviewer activation checklist for branch protection.

### Exit evidence

- Updated governance docs in `docs/`.
- Successful run of `python scripts/check_assumption_release_guardrails.py`.
- Reviewed branch-protection transition decision record.

## Phase 2 (weeks 5-10): actuarial fidelity upgrades

### Deliverables

- Cohort-aware portfolio scenario design and implementation.
- Dynamic lapse v2 design package and initial implementation.
- Scenario-set governance catalog design.

### Exit evidence

- New parity/integration tests for cohort behavior.
- Lapse v2 invariant/property tests.
- Updated product/portfolio specs documenting new behavior.

## Phase 3 (weeks 11-18): platform hardening

### Deliverables

- Incremental `pricing_ui.py` decomposition into feature modules.
- Canonical product-definition schema with migration adapters.
- Durable run ledger MVP wired to CLI first.

### Exit evidence

- Reduced `pricing_ui.py` responsibility with equivalent test coverage.
- Passing invariant tests with simplified registry synchronization burden.
- Sample reproducible run audit record from ledger.

## Phase 4 (weeks 19-22): operationalization and API

### Deliverables

- Internal API MVP for pricing/portfolio orchestration.
- Ongoing monitoring KRI jobs and reporting hooks.
- Release process integration for governance evidence.

### Exit evidence

- API smoke tests and auth/control design note.
- KRI report sample covering parity and benchmark drift.
- Release checklist updated with governance evidence capture.

## Success metrics

- **Traceability:** 100% of release runs identify assumption set and model version.
- **Governance:** all Tier 1 changes have explicit reviewer/challenger evidence.
- **Fidelity:** cohort-aware portfolio results show reduced approximation risk on heterogeneous books.
- **Maintainability:** new product onboarding touches fewer manual registry surfaces.
- **Operational control:** post-merge monitoring detects drift before release.

## Artifact map

| Category | Artifact | Owner |
|---|---|---|
| Governance | `docs/model_inventory.md` | Model governance lead |
| Governance | `docs/assumption_governance.md` | Assumption steward |
| Governance | `scripts/check_assumption_release_guardrails.py` logs | Release owner |
| Actuarial | parity + benchmark test reports | Product actuary |
| Platform | migration design notes + test evidence | Platform engineering |
| Operations | KRI monthly report | Model risk governance |
