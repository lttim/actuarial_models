# UI decomposition — migration map

`pricing_ui.py` is currently 4,118 LOC across 49 top-level functions. It is
also the only place where Streamlit `st.session_state` keys are minted, so a
naive split would silently break what-if/excel-replicator state plumbing.

This package (`annuity_model/ui/`) is the **target** structure for the
decomposition planned in Phase 2 of the hardening roadmap. The goal is to
move the existing functions here progressively, **one logical page at a
time**, with full Streamlit-test coverage at each step.

## Target layout

```
annuity_model/
└── ui/
    ├── app.py                  # ≤ 400 LOC: nav, sidebar, top-level orchestration
    ├── pages/
    │   ├── overview.py         # product picker + landing
    │   ├── pricing_run.py      # _render_pricing_run, run_form binding
    │   ├── what_if.py          # _render_what_if + scenario UI
    │   ├── excel_replicator.py # _render_excel_replicator export-then-recompute
    │   ├── alm.py              # _render_alm + projection visualisation
    │   └── unit_tests.py       # in-app pytest summary
    ├── widgets/                # _render_* helpers shared across pages
    └── forms/
        └── run_form_state.py   # rename of existing pricing_run_form_state.py
```

## Move map (current → target)

| Current location                          | Target                            |
|-------------------------------------------|-----------------------------------|
| `pricing_ui.py` :: `_render_overview`     | `ui/pages/overview.py`            |
| `pricing_ui.py` :: `_render_pricing_run`  | `ui/pages/pricing_run.py`         |
| `pricing_ui.py` :: `_render_what_if`      | `ui/pages/what_if.py`             |
| `pricing_ui.py` :: `_render_excel_replicator` | `ui/pages/excel_replicator.py`|
| `pricing_ui.py` :: `_render_alm`          | `ui/pages/alm.py`                 |
| `pricing_ui.py` :: `_render_unit_tests`   | `ui/pages/unit_tests.py`          |
| `pricing_ui.py` :: helpers (`_render_*`)  | `ui/widgets/<feature>.py`         |
| `pricing_run_form_state.py`               | `ui/forms/run_form_state.py`      |
| `pricing_ui.py` :: `main`, sidebar nav    | `ui/app.py`                       |

## Migration rules (do not skip)

1. **One page per PR.** Move one logical page; keep `pricing_ui.py` importing
   that page from the new location for backward compatibility (one-line
   alias) until the entire move lands.
2. **Public surface only.** Pages MUST import from
   `annuity_model` (the package, see `__init__.py`), never from individual
   modules. This is what kills the current cross-module coupling.
3. **`session_state` discipline.** Every key minted by a page must be
   prefixed with the page name (e.g. `pricing_run__contract_type`). The
   first PR of the migration adds a session-state audit script under
   `scripts/`.
4. **Behaviour test first.** Add a Streamlit-tester or selector-level test
   for the page **before** moving it; the test should pass against
   `pricing_ui.py` first, then unchanged after the move.
5. **No new logic in `pricing_ui.py`.** All new pages / widgets land here.

## Why phased

The audit (Phase 2 plan, `architecture explore` report) found `session_state`
keys minted at ~30 distinct sites, often shared across pages. A single big
split would invariably mis-route at least one of them and produce a
confusing UI bug days later. One page per PR with a behaviour test before
each move keeps blast radius small.
