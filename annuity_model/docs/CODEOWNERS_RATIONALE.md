# CODEOWNERS rationale

`/.github/CODEOWNERS` enumerates the files that require an owner review
before merge. This document explains *why* each block exists, so future
maintainers don't accidentally weaken a load-bearing gate.

## Parity-critical actuarial surface

| Path                                          | Why owner review is required                                              |
|-----------------------------------------------|---------------------------------------------------------------------------|
| `annuity_model/src/annuity_model/pricing_projection.py`         | Source of truth for SPIA cashflows + ALM. A silent change here is a release-stopping defect. |
| `annuity_model/src/annuity_model/term_projection.py`            | Term Life liability cashflow definition; consumed by Excel builder.       |
| `annuity_model/src/annuity_model/rila_projection.py`            | RILA crediting / cap / floor logic; consumed by Excel builder.            |
| `annuity_model/src/annuity_model/alm_excel_ladder.py`           | Generates the Excel formulas that must match Python -- any divergence breaks parity. |
| `annuity_model/src/annuity_model/excel_workbook_validator.py`   | The static gate that catches malformed formulas. Loosening it loses safety. |
| `annuity_model/src/annuity_model/product_registry.py`           | The dispatch surface that lets a new product land in 2 files; structural correctness matters here more than perf. |
| `annuity_model/src/annuity_model/parity_constants.py`           | The immutable tolerance contract. Every value here is referenced by code, tests, and rendered docs. |

## Excel builders

| Path                                          | Why                                                                       |
|-----------------------------------------------|---------------------------------------------------------------------------|
| `annuity_model/src/annuity_model/build_pricing_excel_workbook.py` (SPIA) | One off-by-one in the liability column letter and ModelCheck explodes.    |
| `annuity_model/src/annuity_model/build_term_excel_workbook.py`   | Same.                                                                     |
| `annuity_model/src/annuity_model/build_rila_excel_workbook.py`   | Same; plus RILA must use column M, not S.                                 |
| `annuity_model/src/annuity_model/recalc_excel_shared.py`         | Recalc helper used by parity tests; bug here masks formula errors.        |
| `annuity_model/src/annuity_model/product_excel.py`               | Dispatcher across per-product builders (all implemented products).        |

## Load-bearing internals (added P0 hardening 2026-04)

These were previously covered only by the default `* @lttim` line. Listed
explicitly so reviewers see them flagged as parity/UI critical:

| Path                                          | Why                                                                       |
|-----------------------------------------------|---------------------------------------------------------------------------|
| `annuity_model/src/annuity_model/liability_layouts.py`           | Single source of truth for liability column letters per product. RILA `M` vs SPIA/Term `S` lives here; a typo breaks every workbook. |
| `annuity_model/src/annuity_model/liability_dispatch.py`          | Pricing-result -> LiabilityPath converter registry. Renaming a result class without updating the registered key silently breaks ALM. |
| `annuity_model/src/annuity_model/data_registry.py`               | Versioned data + sha256 lock for mortality/yield CSVs. A silent dataset bump invalidates parity. |
| `annuity_model/src/annuity_model/pricing_ui.py`                  | Monolithic Streamlit UI. Bug class includes silently dropping widget values into hard-coded contract fields (see Term wiring fix). |
| `streamlit_app.py`                                               | App entry point; reroutes / launcher logic.                               |
| `annuity_model/src/annuity_model/pricing_run_form_state.py`      | Run-form session-state defaults & numeric input wrapping. Affects every product's UI inputs. |
| `annuity_model/src/annuity_model/_logging.py`                    | Structured logging configuration; controls what reaches operators.        |
| `annuity_model/src/annuity_model/_observability.py`              | OpenTelemetry tracing decorator. P3 wiring will route engine entry points through this. |

## Parity contracts and tests

| Path                                          | Why                                                                       |
|-----------------------------------------------|---------------------------------------------------------------------------|
| `docs/model_parity_contract.md`               | Public guarantee. Tolerance change must be signed off.                    |
| `docs/rila_parity_contract.md`                | Same.                                                                     |
| `docs/parity_test_checklist.md`               | Release checklist; numbers must match the contract.                       |
| `tests/parity/`                               | The gate. Weakening any tolerance here without an owner is a regression.  |

## Agent guidance

| Path                                          | Why                                                                       |
|-----------------------------------------------|---------------------------------------------------------------------------|
| `AGENTS.md` (root)                            | Load-bearing for every Cursor / coding-agent session.                     |
| `annuity_model/AGENTS.md`                     | Same, scoped to the package.                                              |
| `annuity_model/docs/AI_AGENT_TEAM_PROTOCOL.md` | Defines autonomous multi-agent authority, staffing, and evidence rules.   |
| `annuity_model/scripts/agent_preflight.py`    | Writes Team Run Packets and may run selected completion gates.            |
| `annuity_model/scripts/agent_team_router.py`  | Selects specialist roles and gates for autonomous agent work.             |
| `annuity_model/.cursor/`                      | The `.mdc` rules that constrain agent behaviour around parity and Excel.  |
| `annuity_model/.cursorrules`                  | Legacy file, still consulted by some agents.                              |

## Control plane

| Path                                          | Why                                                                       |
|-----------------------------------------------|---------------------------------------------------------------------------|
| `.github/workflows/`                          | Skipping CI hooks would let bad commits land on `main`.                   |
| `.pre-commit-config.yaml`                     | Local mirror of CI; weakening it lets contributors ship locally-broken PRs. |
| `annuity_model/pyproject.toml`                | Tool config + version + Python floor; impacts every contributor.          |
| `annuity_model/requirements*.txt`             | Runtime / dev dependency pins; a stealth bump can change numerical output. |
| `annuity_model/requirements.lock`             | Frozen transitive deps; bytewise reproducibility for parity.              |

## How to add a new owned area

1. Decide whether the new file is parity-impacting, builder-impacting,
   governance-impacting, or control-plane-impacting.
2. Add a CODEOWNERS line under the matching block (preserve the
   block ordering).
3. Add a row to the appropriate table here with a one-sentence rationale.
4. If you are adding a new product, also add the new builder and the new
   parity test path.

## Second-CODEOWNER upgrade path

The repo currently has exactly one CODEOWNER (`@lttim`). GitHub blocks
self-approval on protected branches, so requiring `>= 1` reviewer in
`.github/branch-protection.json` would deadlock every PR. The active
profile therefore sets `required_pull_request_reviews = null`. This is a
*deferred* gate, not a missing one: every piece of infrastructure
needed to flip it on is already committed. To activate:

1. **Onboard the second reviewer.** Either an individual maintainer with
   actuarial signoff authority, or a GitHub team. The placeholder
   handle this repo references is `@lttim/actuarial-reviewers`; if you
   create the team under a different handle, do a one-shot
   find-and-replace across `.github/CODEOWNERS`,
   `.github/branch-protection.with-second-reviewer.json`, this file,
   and `annuity_model/docs/CHANGELOG.md`.
2. **Uncomment the CODEOWNERS hooks.** Every parity-critical override
   in `.github/CODEOWNERS` carries a `TODO(second-owner):` comment with
   the literal line to uncomment. Drop the `# TODO(second-owner):`
   prefix on each one in a single PR.
3. **Swap the branch-protection profile.** Apply
   `.github/branch-protection.with-second-reviewer.json` instead of the
   current `.github/branch-protection.json`:

   ```bash
   gh api -X PUT repos/:owner/:repo/branches/main/protection \
     --input .github/branch-protection.with-second-reviewer.json
   ```

   The two profiles differ ONLY in
   `required_pull_request_reviews` -- the status-checks list, linear
   history requirement, force-push policy, and conversation-resolution
   gate are byte-identical between the two files.
   `tests/test_branch_protection_drift.py` enforces that. If you ever
   add a status check or change a flag, update both files in the same
   PR or the drift test will fail.
4. **Promote the deferred file to active.** Once the second-reviewer
   profile has been live for one full release cycle without rollback,
   replace the contents of `.github/branch-protection.json` with the
   contents of `.github/branch-protection.with-second-reviewer.json`
   and delete the latter. The drift test will be removed in the same
   PR.
