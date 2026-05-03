# Agent Instructions — SPIA Annuity Model

This project implements SPIA, Term Life, RILA, MYGA, FIA, VA, WL, UL, IUL, and VUL
pricing/projection with ALM projection and two synchronised calculation engines where
applicable: Python and Excel. RILA and IUL are classified as mechanics-production
product-mechanics prototypes: policy mechanics and selected-path Excel replication are
production-grade, while assumption governance remains advisory until Actuary SME signoff.

For a cross-platform overview of project-development controls (rules, skills, delegated
reviews/subagents, regressions, and governance workflow), see
`../PROJECT_DEVELOPMENT_GUIDE.md`.

> **AI coding agents:** read [`docs/AI_AGENT_PREFLIGHT.md`](docs/AI_AGENT_PREFLIGHT.md)
> *first*. It contains the source-of-truth map, decision tree, and hard
> rules an autonomous agent must follow on this repo. This file
> (`AGENTS.md`) remains the canonical owner of the four-gate list below;
> the pre-flight doc just routes agents to the right runbook before they
> get there.

## Non-negotiable parity requirement

Every code change must maintain zero-discrepancy between the Python ALM engine and the
generated Excel workbook. Read and follow `docs/model_parity_contract.md` before modifying
any calculation logic.

## Before completing any task -- canonical gates

> This block is the *single source of truth* for the "what do I run before
> claiming a task is done?" question. Any other doc that needs to state these
> gates (root `AGENTS.md`, `actuarial-parity.mdc`, `parity_test_checklist.md`,
> `CONTRIBUTING.md`) MUST link here instead of restating.

```bash
# 1. Parity gate (blocks any merge on failure; includes the SME lite golden)
python -m pytest tests/parity -q

# 2. Full unit-test gate
python -m pytest -q

# 3. End-to-end smoke (all implemented products + full Excel validator)
python scripts/deep_smoke.py

# 4. Tolerance contract is in sync with parity_constants.py
python scripts/render_parity_contract.py --check
```

All four must exit 0.

### Ring 7 -- portfolio acceptance (optional superset)

For changes touching **portfolio** aggregation, inforce I/O, portfolio Excel,
or the portfolio CLI/UI (`ANNUITY_MODEL_PORTFOLIO_V1` surfaces), also run from
the **repository root** (not only `annuity_model/`):

```bash
just portfolio-acceptance
```

That recipe runs the four gates above, then `tests/parity/portfolio`,
`tests/integration`, portfolio-enabled `deep_smoke`, an explicit
`render_parity_contract.py --check`, the CLI JSON golden, and **`just
actuary-review-full`** (Gate 5 evidence, full scope). CI enforces the same
bundle via the **portfolio acceptance (ring 7)** workflow when branch
protection lists that required check. Details:
[`docs/runbooks/portfolio_run.md`](docs/runbooks/portfolio_run.md).

### Gate 5: Actuary SME review (recursive)

After the four canonical gates exit 0, the **Actuary SME review** is a
mandatory fifth gate when the session edited any file in the
CALCULATION or TOLERANCE branches of
[`docs/AI_AGENT_PREFLIGHT.md`](docs/AI_AGENT_PREFLIGHT.md). It is a
**recursive** gate: rather than a single command, it runs an
autonomous fix-and-rereview loop, defined in full at
[`.cursor/rules/actuary-sme-protocol.mdc`](../.cursor/rules/actuary-sme-protocol.mdc).

Trigger forms (all equivalent):

- Explicit command: `!actuaryreview` (with optional `full`,
  `<product>`, or `status` argument).
- Natural language: any phrasing containing "actuary review", "have
  the actuary review", "ask the actuary SME", "actuarial review
  please", etc. The rule routes these through the same orchestration.
- Auto-fired: the rule self-fires when the session edited files
  matching the auto-trigger globs (engines, builders, parity
  constants, actuarial benchmarks, product subpackages).

Termination conditions:

- **APPROVE** (or APPROVE-WITH-NOTES with no `[AGENT-FIXABLE]`
  items): the loop exits cleanly and the task may complete.
- **BLOCK** with `[AGENT-FIXABLE]` items: the parent agent applies
  the fixes, re-runs gate 1 (parity), and re-invokes the SME --
  iterating up to `MAX_ITERATIONS` (default 5).
- **Escalation** (max iterations exceeded, the same finding recurs,
  or any `[NEEDS-HUMAN-JUDGMENT]` finding inside a `BLOCK` verdict):
  the loop stops, prints a chain of verdict files under
  `.cursor/actuary-reviews/`, and the task does NOT claim complete.
  The user is **not** prompted for a next step at any point.

Verdict files are stored at
`.cursor/actuary-reviews/iter-<N>-<UTC>-<scope>.md` (gitignored, like
`.cursor/handoffs/`). The skill that defines the SME persona,
checklist, and YAML-frontmatter verdict template is at
[`.cursor/skills/actuary-sme/SKILL.md`](.cursor/skills/actuary-sme/SKILL.md).

Two always-on gates inside gate (2) above are worth calling out explicitly
because they catch the bug classes the parity engine cannot:

* **`tests/ui/test_apptest_full_workflow.py`** — runs Streamlit's
  `AppTest` harness end-to-end against `pricing_ui.py` AND
  `streamlit_app.py` (the Streamlit Cloud entry). Asserts every
  sidebar section renders without exceptions, that clicking
  "Run pricing" populates `st.session_state['pricing_res']` for
  every implemented product, and that the Excel download bytes pass
  strict `excel_workbook_validator`. If you touch `pricing_ui.py`,
  any product adapter, or the workbook builder dispatch, this gate
  is your first signal.
* **`tests/parity/test_excel_recalc_per_product.py`** — per-product
  Excel↔Python workbook gate. The "always-on" layer asserts the
  Python literals the builder bakes into ModelCheck column B equal
  the engine within `MODELCHECK_TOL` for every product (no skip);
  the formula-contract layer asserts ModelCheck column C links to the
  canonical validated liability summary rows. The pre-existing
  `tests/parity/test_runtime_excel_recalc.py` is retained as a focused
  SPIA ModelCheck wiring regression.

If a task changes Excel-generating code, ALSO validate the regenerated
workbook and inspect the `ModelCheck` sheet links. The parity tests assert
Python snapshots and formula wiring, while the static validator catches
formula syntax and cross-sheet reference issues. Step-by-step recipe in
[docs/runbooks/regenerate_excel_cache.md](docs/runbooks/regenerate_excel_cache.md).

If parity fails, follow
[docs/runbooks/investigate_parity_break.md](docs/runbooks/investigate_parity_break.md).
If the validator fails, follow
[docs/runbooks/debug_validator_failure.md](docs/runbooks/debug_validator_failure.md).
**Never widen a tolerance to make a test pass.** Tolerance changes route
through `parity_constants.py` plus `model_change_log.md`.

## Key files

| File | Purpose |
|------|---------|
| `pricing_projection.py` | Python ALM engine (source of truth for calculation logic) |
| `alm_excel_ladder.py` | Excel formula generator — must match Python exactly |
| `build_pricing_excel_workbook.py` | Workbook builder and OOXML cache injector |
| `docs/model_parity_contract.md` | Parity contract: tolerances, tie-break, epsilon policies |
| `docs/parity_test_checklist.md` | Release gate checklist |
| `tests/parity/test_alm_parity.py` | Parity regression tests |
| `rila_projection.py` | RILA liability / crediting (Python) |
| `build_rila_excel_workbook.py` | RILA Excel workbook + `ModelCheck` |
| `docs/rila_product_spec.md` | RILA mechanics-production product definition |
| `docs/iul_product_spec.md` | IUL mechanics-production product definition |
| `docs/rila_parity_contract.md` | RILA Python ↔ Excel parity addendum |
| `tests/parity/test_rila_parity.py` | RILA parity tests |
| `portfolio_runner.py` | Multi-policy pricing loop + optional process pool |
| `liability_aggregation.py` | Union-grid `LiabilityPath` sums (total + by type) |
| `build_portfolio_excel_workbook.py` | Portfolio rollup workbook + `ModelCheck` |
| `inforce_io.py` / `inforce_parsers.py` | Inforce CSV / Excel → `PolicyInput` |
| `docs/portfolio_runner_spec.md` | Portfolio v1 product / JSON / Excel spec |
| `docs/portfolio_parity_contract.md` | Portfolio parity addendum |

## Critical rules

1. Any change to disinvestment/reinvestment ordering logic requires a new parity test.
2. Never change epsilon values without updating the contract and adding a boundary test.
3. Never rely on raw floating-point comparison of `t_rem` values for ordering — always use epsilon tie-breaking.
4. Every bug fixed must have a permanent regression test capturing the exact scenario.
5. Step-level reconciliation (monthly state) not just final surplus.
6. **Excel formulas must pass static validation before saving.** Every workbook builder must
   call `excel_workbook_validator.validate_workbook_or_raise(wb)` immediately before
   `wb.save(...)`. The validator now checks both *syntax* and *cross-sheet semantics*:
   - Balanced parentheses and string quotes (catches f-string concatenation bugs).
   - Correct argument counts for every known function (no `IF(cond, value)` without an
     explicit false branch — Excel will repair the file with "Removed Records: Formula
     from /xl/worksheets/sheetN.xml part" and replace cells with `#DIV/0!`).
   - No embedded Excel error literals (`#REF!`, `#NAME?`, `#DIV/0!`, ...) inside formula
     bodies.
   - No trailing empty arguments (`IF(a, b, )` / `IFERROR(x, )`) — those are almost
     always an f-string that lost its substitution. Write `,""` or `,0` explicitly.
   - **Cross-sheet column resolution.** Every `Sheet!Col` reference *and* every
     `Sheet!Col` literal embedded in `INDIRECT(...)` must resolve to a column that
     actually has data on the target sheet. This catches the class of silent
     reconciliation bug where a SPIA-style `Liabilities!S` reference is generated
     from a RILA workbook (RILA puts `ExpTotalCF` in column M, not S). Excel
     coerces the missing column to zero in `SUMPRODUCT`/`INDEX`, so without this
     check the failure only shows up as drift in `ModelCheck` after Excel recalcs.
   When introducing a new built-in function in any builder, register its arity in
   `excel_workbook_validator.FUNCTION_ARITIES`. When introducing a new product, the
   shared ALM helper `_write_alm_projection_sheet` accepts `liability_total_col`
   (default `"S"` for SPIA/Term) and `liability_discount_col` (default `"O"`) — pass
   the column letters that match your product's `Liabilities` layout (e.g. RILA
   uses `liability_total_col="M"`). The end-to-end gate
   `tests/test_excel_export_validation.py` builds workbooks for every implemented
   workbook-backed product, including RILA and IUL formula-mechanics sheets, and
   runs the validator on them; that test must always pass.

## Parity kit for future products

See `../actuarial_parity_kit/` for a reusable governance template to carry forward to new
actuarial product repos. Copy that directory into any new repo and adapt the test fixtures.
