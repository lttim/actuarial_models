# Handoff — `annuity_model` (multi-product)

Short snapshot for humans. For structured AI session continuity, use `!handoff`
/ `!recall` per `.cursor/rules/handoff-recall.mdc` (files under `.cursor/handoffs/`).

## Tip of `main`

Run at repo root: `git log -1 --oneline`

## Cross-platform setup

- **macOS / Linux:** `cd annuity_model && ./bootstrap_macos.sh`
- **Windows:** `cd annuity_model && bootstrap.bat`, or the PowerShell flow in root `AGENTS.md`
- **venv:** `annuity_model/.venv/` (never a nested `.git` under `annuity_model/`)

## Verification gates (must all be green before merge)

From `annuity_model/` with `.venv` active — see **`AGENTS.md`** in this directory
for the authoritative four-command block. Typical full suite:

```bash
.venv/bin/python3 -m pytest tests/ tests/parity/ -q
```

Last local reference run: **907 passed, 19 skipped** (exit 0). Counts change as
tests are added; **exit code 0** is the invariant.

## Product surface

Ten pricing engines + Excel builders behind `ProductRegistry` (SPIA, Term,
RILA, MYGA, FIA, VA, WL, UL, IUL, VUL). Narrative plan:
`docs/seven_product_rollout_plan.md`.

## Parity / release (do not skip)

- `docs/model_parity_contract.md` — SPIA/ALM (generated tables)
- `docs/rila_parity_contract.md` — RILA addendum (generated)
- `docs/portfolio_parity_contract.md` — portfolio aggregation addendum
- `docs/rila_product_spec.md`, `docs/portfolio_runner_spec.md` — specs
- Every `wb.save(...)`: **`validate_workbook_or_raise(wb)`** immediately before
- RILA `liability_total_col="M"`; SPIA/Term (and most others) use `"S"` per
  `liability_layouts.py`

## Optional next steps

- See `docs/actuarial_fidelity_backlog.md` and `docs/platform_engineering_roadmap.md`
  for staged work; `docs/index.md` for the full doc map.

## Open issues

None recorded in this file. If validator or parity fails, use
`docs/runbooks/debug_validator_failure.md` and
`docs/runbooks/investigate_parity_break.md`.
