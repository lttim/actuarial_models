# AI Agent Pre-flight

**Audience:** *AI coding agents* (Cursor, Claude Code, Codex, etc.) operating
on this repository. Human contributors should start with root
[`AGENTS.md`](../../AGENTS.md) and [`CONTRIBUTING.md`](../../CONTRIBUTING.md),
then [`AGENTS.md`](../AGENTS.md) for the canonical four-gate commands,
[`Justfile`](../../Justfile) recipes (`just preflight`, `just docs-check`),
and the per-task runbooks under [`docs/runbooks/`](runbooks/).

This file is the **single entry point** an autonomous agent should read
*before* proposing any code change. It does four things, in order:

1. Tells the agent **which canonical doc owns each rule**, so there is no
   ambiguity about which file is the source of truth.
2. Gives a **decision tree** for the most common task classes
   ("change a calculation", "add a product", "change a tolerance",
   "modify the UI", ...). Each leaf points to the runbook *and* the
   gates that must pass.
3. Enumerates the **four canonical gates** the agent must run (or arrange
   to be run) before claiming a task is done -- mirrored verbatim from
   `AGENTS.md` so the agent never has to fetch a second file just to
   know what to run.
4. Routes multi-agent work through
   [`AI_AGENT_TEAM_PROTOCOL.md`](AI_AGENT_TEAM_PROTOCOL.md),
   `scripts/agent_team_router.py`, and `scripts/agent_preflight.py` so
   validators, builders, reviewers, and docs stewards are staffed
   automatically when the task warrants it.

If anything in this doc disagrees with `AGENTS.md`, **`AGENTS.md` wins**
(it is the human-facing source of truth and the one the parity-kit
template is checked against in `tests/test_kit_template_parity.py`).
File a fix here, do not silently diverge.

---

## 0. Source-of-truth map

Before changing *anything*, know which doc owns the rule you are about
to bend.

| Concern | Source of truth | Notes |
|---|---|---|
| Four canonical gates ("what do I run?") | [`AGENTS.md`](../AGENTS.md) -- "Before completing any task" | Mirrored in §3 below for convenience. |
| Numerical parity tolerances | [`parity_constants.py`](../src/annuity_model/parity_constants.py) | Every change must also append to [`docs/model_change_log.md`](model_change_log.md). CI (`parity-constants-log-guard`) enforces this. |
| Tolerance contract narrative (the *why*) | [`docs/model_parity_contract.md`](model_parity_contract.md) | Auto-rendered from `parity_constants.py` via `scripts/render_parity_contract.py`. Never edit by hand. |
| Per-product Excel column layout | [`liability_layouts.py`](../src/annuity_model/liability_layouts.py) | Every product must have an entry; cross-sheet validator enforces. |
| ProductType enum & adapter Protocol | [`product_registry.py`](../src/annuity_model/product_registry.py) | Enum, adapter seed implementations, term label parsers, and `validate_run_inputs(state)` compatibility hook. |
| Per-product source of truth (engine/excel/ui/converter/formatter/capability/mortality/validator/order) | [`products/<name>/__init__.py`](../src/annuity_model/products/) | Use `register_product(ProductDefinition(...))`. Legacy public registry views derive from this canonical record. |
| Streamlit run-form state keys | [`pricing_run_form_state.py`](../src/annuity_model/pricing_run_form_state.py) | Constants live here -- never hardcode `"run_*"` literals elsewhere. |
| Multi-policy portfolio runner | [`portfolio_runner_spec.md`](portfolio_runner_spec.md) + [`portfolio_runner.py`](../src/annuity_model/portfolio_runner.py) | UI/CLI enablement: [`portfolio_config.py`](../src/annuity_model/portfolio_config.py) (`portfolio_v1_enabled`); see also `docs/portfolio_parity_contract.md`. |
| ALM dispatch (pricing-result -> liability-path) | [`liability_dispatch.py`](../src/annuity_model/liability_dispatch.py) | Plug-in registry; `register_liability_path_converter`. |
| Workbook-builder dispatch | [`product_excel.py`](../src/annuity_model/product_excel.py) | Plug-in registry; `register_builder`. |
| Static Excel-formula validator | [`excel_workbook_validator.py`](../src/annuity_model/excel_workbook_validator.py) | All builders MUST call `validate_workbook_or_raise(wb)` before `wb.save(...)`. |
| Autonomous AI team staffing | [`docs/AI_AGENT_TEAM_PROTOCOL.md`](AI_AGENT_TEAM_PROTOCOL.md) + [`scripts/agent_team_router.py`](../scripts/agent_team_router.py) | The orchestrator automatically selects core/dynamic specialist roles from changed files and task traits. |
| Team Run Packet / preflight evidence | [`scripts/agent_preflight.py`](../scripts/agent_preflight.py) + [`scripts/check_team_run_packet_evidence.py`](../scripts/check_team_run_packet_evidence.py) | Writes `.agent-team-runs/` packets, optionally runs selected gates, and blocks broad/high-risk changes without completed staffing/gate/signoff evidence. |
| Adding a new product (full walkthrough) | [`README.md`](../README.md) -- "Adding a new product" | The `scripts/scaffold_product.py` CLI generates step 4's boilerplate. |
| CI workflow definitions | [`.github/workflows/`](../../.github/workflows/) | `parity-gate.yml` is always-on and listed in branch protection. |
| Code-ownership routing | [`.github/CODEOWNERS`](../../.github/CODEOWNERS) and [`docs/CODEOWNERS_RATIONALE.md`](CODEOWNERS_RATIONALE.md) | Editing parity-critical files requires the listed reviewer. |

---

## 1. Decision tree

Pick the **lowest** matching node and follow only that branch. If two
nodes appear to apply, pick the more restrictive (calculation > UI > docs).

Before implementation, run the team router/preflight so staffing is automatic
and evidence is captured:

```bash
python scripts/agent_preflight.py --objective "<task>" --write-packet
```

Before completion on the final diff, add `--run-gates` when the selected gate
set is appropriate for the task and environment. The router does not replace
the branch rules below; it staffs the agents and records the evidence needed to
prove they were followed.

For broad or high-risk changes, the final packet is not optional. The local
pre-commit hook and PR CI run `scripts/check_team_run_packet_evidence.py`.
Locally, it validates the completed `.agent-team-runs/` packet for the staged
diff. In CI, ignored packet files are unavailable, so the PR body must include
a `Team Run Packet Evidence` excerpt with selected roles, validation gates,
review findings, unresolved risks, and a `COMPLETE` final signoff.
If a gate is intentionally deferred because a later staged chunk will run a
larger regression sweep, the packet JSON must record that gate under
`deferred_gate_results` with both `reason` and `next_validation`; silent gate
omissions are treated as incomplete evidence.

```
START
│
├── Is the task documentation-only? (typo, comment, README polish)
│      → branch: DOC-ONLY
│
├── Does the change touch ANY of:
│       src/annuity_model/parity_constants.py, model_parity_contract.md,
│       parity_test_checklist.md, MODELCHECK_TOL,
│       any tolerance constant?
│      → branch: TOLERANCE
│
├── Does the change touch ANY of:
│       src/annuity_model/pricing_projection.py, src/annuity_model/term_projection.py, src/annuity_model/rila_projection.py,
│       src/annuity_model/alm_excel_ladder.py, src/annuity_model/build_*_excel_workbook.py,
│       src/annuity_model/excel_builder_helpers.py, src/annuity_model/excel_workbook_validator.py,
│       src/annuity_model/liability_layouts.py, src/annuity_model/liability_dispatch.py,
│       src/annuity_model/product_registry.py, src/annuity_model/product_excel.py,
│       src/annuity_model/products/<name>/(engine|excel|schema).py?
│      → branch: CALCULATION
│
├── Is the task "add a new product"?
│      → branch: NEW-PRODUCT
│
├── Does the change touch portfolio / inforce / multi-policy aggregation?
│       (src/annuity_model/portfolio*.py, src/annuity_model/liability_aggregation.py, src/annuity_model/inforce_io.py,
│        src/annuity_model/inforce_parsers.py, src/annuity_model/build_portfolio_excel_workbook.py,
│        src/annuity_model/products/*/inforce.py)
│      → branch: PORTFOLIO
│
├── Does the change touch the Streamlit UI?
│       (src/annuity_model/pricing_ui.py, streamlit_app.py, src/annuity_model/pricing_run_form_state.py,
│        src/annuity_model/ui/, src/annuity_model/products/<name>/ui.py)
│      → branch: UI
│
├── Does the change touch CI / pre-commit / Justfile / branch-protection?
│      → branch: INFRA
│
└── (default)  → branch: SUPPORT  (tests-only, scripts, docstrings,
                                    refactor with no behavior change)
```

### Branch: DOC-ONLY

Required gates: docs lint only.

```bash
just docs-check     # markdownlint + link-check + rendered-contract diff
```

You do NOT need to run pytest unless you touched a `.py` file.
**However**, if you edited `AGENTS.md` you MUST also run
`tests/test_kit_template_parity.py` -- the parity kit template is held in
lock-step with the canonical AGENTS.md by that test.

### Branch: TOLERANCE

This is the highest-risk change class in the repo.

1. Edit only `src/annuity_model/parity_constants.py`. Never widen a tolerance to make a
   test pass.
2. Append a dated entry to [`docs/model_change_log.md`](model_change_log.md)
   describing the change, the parity scenario that exposed it, and the
   reviewer who approved. CI gate `parity-constants-log-guard` will
   reject the PR otherwise.
3. Re-render the parity contract: `python scripts/render_parity_contract.py`.
4. Run **all four canonical gates** (§3). Then add a boundary regression
   test under `tests/parity/` capturing the exact scenario.
5. **Trigger the Actuary SME review** (`!actuaryreview`, or any
   natural-language "actuary review" request) and obtain a clean
   APPROVE before declaring the task done. Tolerance changes are the
   highest-risk class precisely because they can mask methodology
   drift; the SME's autonomous fix-and-rereview loop catches that
   class. See §3.5 for the loop's termination semantics.
6. Open the PR with the parity-contract diff and the Actuary SME
   verdict path in the description.

Runbook: [`docs/runbooks/investigate_parity_break.md`](runbooks/investigate_parity_break.md).

### Branch: CALCULATION

This is what the parity gates exist for.

1. Make the Python change first (engine source-of-truth).
2. Mirror the change in the Excel formula generator (`alm_excel_ladder.py`
   for SPIA/Term, the relevant `build_*_excel_workbook.py` for RILA, etc.).
3. Run the four canonical gates (§3). Pay particular attention to gate 1
   (parity), the static workbook validator, and ModelCheck formula-link
   checks.
4. If gate 1 fails, follow [`docs/runbooks/investigate_parity_break.md`](runbooks/investigate_parity_break.md).
   If the validator fails, follow
   [`docs/runbooks/debug_validator_failure.md`](runbooks/debug_validator_failure.md).
5. Add a `@pytest.mark.regression` test capturing the bug or new behavior.
6. **Trigger the Actuary SME review** (`!actuaryreview`, or any
   natural-language "actuary review" request). Calculation changes
   are this branch's whole reason for existing; the SME catches the
   "internally consistent but actuarially nonsense" class of bug that
   parity tests by construction cannot. See §3.5.

### Branch: NEW-PRODUCT

Use the scaffolder, then fill in the actuarial code.

```bash
python scripts/scaffold_product.py \
    --code <name> \
    --display-name "<Human Name>" \
    --contract-class <Name>Contract \
    --result-class <Name>ProjectionResult
```

The script generates the `src/annuity_model/products/<name>/{__init__,schema,engine,excel,ui}.py`
shims. The follow-up checklist printed by the script lists the manual
steps (engine implementation, ProductType enum member, dispatch converter,
liability layout, Excel builder, parity test).

The meta-invariant tests (`tests/test_meta_invariants.py`,
`tests/test_products_registry.py`, `tests/test_mypy_strict_glob.py`) will
keep failing with helpful messages until every wire is connected --
treat that as your todo list.

Then run the four canonical gates (§3).

### Branch: PORTFOLIO

Read [`docs/portfolio_runner_spec.md`](portfolio_runner_spec.md) and
[`docs/portfolio_parity_contract.md`](portfolio_parity_contract.md).

Required checks (in addition to the four canonical gates when portfolio code
paths change):

```bash
cd annuity_model
ANNUITY_MODEL_PORTFOLIO_V1=1 python -m pytest tests/parity/portfolio -q
ANNUITY_MODEL_PORTFOLIO_V1=1 python -m pytest tests/integration/test_portfolio_cli.py -q
```

Actuary SME: `!actuaryreview portfolio` (or `product:portfolio`) after
calculation-facing edits; use `just portfolio-acceptance` before merge when
touching the end-to-end portfolio surface.

### Branch: UI

Streamlit changes never bypass parity, but they have an extra constraint:

* All `st.session_state` keys MUST come from `pricing_run_form_state.py`
  constants (never hardcode `"run_*"` literals elsewhere -- a grep test
  enforces this).
* Run [`tests/ui/`](../tests/) AppTest smoke tests in addition to the
  full pytest gate.
* If the change adds a widget that becomes a contract input, also add
  it to `validate_run_inputs(state)` in `product_registry.py`.

Then run the four canonical gates (§3).

### Branch: INFRA

CI / pre-commit / branch-protection / Justfile changes.

* If you change `parity-gate.yml`, also update
  `.github/branch-protection.json` `required_status_checks.contexts`
  (the parity gate must remain a required check).
* If you add a new gate, add it to the four-gate list in `AGENTS.md`,
  to the `just preflight` recipe, and update the `REQUIRED_FRAGMENTS`
  list in `tests/test_kit_template_parity.py`.
* If you change the mypy strict surface, update the `pyproject.toml`
  override block AND verify `tests/test_mypy_strict_glob.py` still
  passes (its `LOAD_BEARING_CORE` list is the human-readable mirror).

Then run the four canonical gates (§3).

### Branch: SUPPORT

Tests-only, refactor-only, docstring-only changes that genuinely cannot
move calculation behavior.

* Run gate 2 (full pytest) at minimum.
* If you touched anything imported by an Excel builder, run gate 3
  (`scripts/deep_smoke.py`) too.
* If unsure: run all four canonical gates (§3). They are cheap.

---

## 2. Hard rules an agent MUST never break

These are non-negotiable. Failing any of them is a bug in the *agent*,
not a "trade-off".

1. **Never widen a tolerance to make a test pass.** Tolerance changes
   route through `parity_constants.py` plus `model_change_log.md`.
2. **Never bypass `validate_workbook_or_raise(wb)`.** Every workbook
   builder MUST call it immediately before `wb.save(...)`.
3. **Never edit `docs/model_parity_contract.md` by hand.** It is rendered
   from `parity_constants.py`; edit the constants and re-run the renderer.
4. **Never hardcode `"run_*"` session-state keys.** Use the constants in
   `pricing_run_form_state.py`. A grep test enforces this.
5. **Never add a product without an entry in `liability_layouts.py`.** The
   cross-sheet validator and parity dispatch will silently degrade
   otherwise.
6. **Never silence a meta-invariant test.** Tests under
   `tests/test_meta_invariants.py`, `tests/test_products_registry.py`,
   `tests/test_mypy_strict_glob.py`, and `tests/test_kit_template_parity.py`
   are the *contract*; if they fail, fix the underlying drift, do not
   skip the test.
7. **Never disable the parity gate.** `parity-gate.yml` is always-on and
   listed in branch protection.

---

## 3. Canonical gates (mirrored from AGENTS.md)

> The authoritative copy of this list is the
> "Before completing any task" section of [`AGENTS.md`](../AGENTS.md).
> If anything below diverges from that file, that file wins.

```bash
# 1. Parity gate (blocks any merge on failure)
python -m pytest tests/parity -q

# 2. Full unit-test gate
python -m pytest -q

# 3. End-to-end smoke (all implemented products + full Excel validator)
python scripts/deep_smoke.py

# 4. Tolerance contract is in sync with parity_constants.py
python scripts/render_parity_contract.py --check
```

All four must exit 0. The `just preflight` recipe runs all four
sequentially and prints `READY TO COMMIT` on success; prefer that over
running the four manually.

If a task changes Excel-generating code, ALSO open the regenerated
workbook in Excel (or run `recalc_excel_shared.recalculate_workbook`)
and verify the `ModelCheck` sheet shows 0.00 difference -- the parity
tests load values from openpyxl and CANNOT see formula bugs that only
surface on Excel recalc. Step-by-step recipe in
[`docs/runbooks/regenerate_excel_cache.md`](runbooks/regenerate_excel_cache.md).

---

## 3.5. Gate 5: Actuary SME review (recursive)

After the four canonical gates exit 0, the **Actuary SME** review is a
mandatory fifth gate when the session edited any file in the
CALCULATION or TOLERANCE branches. It is a recursive gate: rather than
a single command, it runs an autonomous fix-and-rereview loop.

Triggers (all equivalent):

- Explicit command: `!actuaryreview` (with optional `full`,
  `<product>`, or `status` argument).
- Natural language: any phrasing containing "actuary review", "have
  the actuary review", "ask the actuary SME", "actuarial review
  please", etc.
- Auto-fired by the always-on rule
  [`.cursor/rules/actuary-sme-protocol.mdc`](../../.cursor/rules/actuary-sme-protocol.mdc)
  when the session edited files matching the auto-trigger globs
  (engines, builders, parity_constants, actuarial_benchmarks, product
  subpackage engines, etc.).

Loop semantics (the rule defines the full state machine; abridged here):

1. The evidence script
   ([`scripts/run_actuary_review.py`](../scripts/run_actuary_review.py))
   collects diff + cached test results + benchmarks into
   `.cursor/actuary-reviews/_evidence-current.md` (overwritten each
   iteration; <5s; reads cached pytest output, does NOT re-run tests).
2. A readonly subagent renders a verdict per
   [`.cursor/skills/actuary-sme/SKILL.md`](../.cursor/skills/actuary-sme/SKILL.md).
   The verdict is YAML frontmatter + markdown body, written to
   `.cursor/actuary-reviews/iter-<N>-<UTC>-<scope>.md`. The subagent is
   created on iter 1 and resumed (`Task(resume=...)`) on iter 2+ to
   preserve context.
3. **Clean APPROVE** (or APPROVE-WITH-NOTES with no `[AGENT-FIXABLE]`
   items): full `just preflight` runs once as the final safety check;
   on green the task may complete.
4. **BLOCK** with `[AGENT-FIXABLE]` items: parent agent applies the
   verbatim Required Actions, runs gate 1 only (parity, ~10s) for
   fast feedback, and re-invokes the SME. Mid-loop full preflight is
   intentionally skipped to keep loop runtime bounded.
5. **Escalation** (any of: `iter >= MAX_ITERATIONS`, recurring
   `[AGENT-FIXABLE]` finding marked unresolved, or any
   `[NEEDS-HUMAN-JUDGMENT]` finding inside a `BLOCK`): loop stops,
   prints the chain of verdict file paths, and the task does NOT
   complete. The user is **not** prompted at any point in the loop.

`MAX_ITERATIONS` defaults to 5; override with the
`ACTUARY_REVIEW_MAX_ITER` environment variable.

---

## 4. When you are stuck

The repo's runbooks are written for exactly this situation. Pick the
one that matches the failure mode:

| Failure | Runbook |
|---|---|
| Parity test went from green to red | [`docs/runbooks/investigate_parity_break.md`](runbooks/investigate_parity_break.md) |
| `excel_workbook_validator` rejected the workbook | [`docs/runbooks/debug_validator_failure.md`](runbooks/debug_validator_failure.md) |
| `ModelCheck` links or workbook validation look wrong | [`docs/runbooks/regenerate_excel_cache.md`](runbooks/regenerate_excel_cache.md) |
| Release / version bump | [`docs/runbooks/release.md`](runbooks/release.md) |
| Streamlit launcher won't double-click | [`docs/runbooks/launcher_double_click.md`](runbooks/launcher_double_click.md) |

If none match, **stop and ask** -- file an issue or surface the question
to the human reviewer. Guessing on a parity-critical surface is a worse
outcome than a slower turnaround.
