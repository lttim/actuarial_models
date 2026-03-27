# Current Working State

- Branch: `main` (synced with `origin/main` after publish; verify with `git status`).
- Workspace: `annuity_model`
- Tip of `main`: `f8d1c24` (published). Feature commit: `e3f1a52` — Streamlit Pricing Run: `pricing_run_form_state` (seed defaults + `run_number_input` + `ensure_session_choice`) fixes min-value first-paint bugs (Term premium, SPIA valuation year / RP-2014 context, other run numerics).
- Prior related commits: Term premium UI seed (`e87e690` area); Excel recalc parity (`acfa999`); Term/SPIA workbook alignment.

# Key Logic Implemented

## Streamlit Pricing Run (`pricing_ui.py` + `pricing_run_form_state.py`)

- **Single source of truth** for initial `run_*` session keys: `build_run_form_seed_defaults` (extend when adding products/inputs).
- **`run_number_input`**: always passes explicit `value=` so `st.number_input` does not default to `value="min"` → wrong UI before first run.
- **`ensure_session_choice`**: before mortality / yield / expense radios when option sets are product-dependent.
- Prior: SPIA Excel recalc strips Dashboard/Runbook; Term `Liabilities` / curves / optional ALM; shared `recalc_excel_shared.py`.

## Tests

- `tests/test_pricing_run_form_state.py` — coercion, seed shape, `ensure_session_choice`.
- Parity + workbook tests unchanged in scope.

## Parity / invariants

- No calculation-engine changes; parity contract unchanged.

# Validation Completed

- `python -m pytest tests/parity/ -v` — 19 passed.
- `python -m pytest tests/ -v` — 111 passed.

# Immediate Next Steps

- None required after publish.

# Unresolved Bugs / Pending Calculations

- No known failing tests.
- Optional: document `run_alm_projection_from_pricing_result` in `AGENTS.md` / parity checklist.
- Optional: apply same `run_number_input` pattern to ALM / What-if tabs if min-value display issues appear there.
