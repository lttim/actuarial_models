# Code_Sandbox — Actuarial Pricing & ALM Platform

Python ↔ Excel parity-tested pricing engine for **SPIA** (single-premium immediate
annuity), **Term Life**, and **RILA** (registered-index-linked annuity) products,
with an ALM (asset–liability management) ladder, a Streamlit pricing UI, and a
static Excel formula validator that gates every workbook before it is written.

## Workspace layout

- [annuity_model/](annuity_model/) — the production package: engines, Excel
  builders, validator, registry, and the Streamlit app.
- [actuarial_parity_kit/](actuarial_parity_kit/) — drop-in templates for spinning
  up a parity-gated actuarial project from scratch (rules, AGENTS template,
  parity contract template, parity-trace exporter).
- [streamlit_app.py](streamlit_app.py) — Streamlit Cloud entry point that wires
  in `annuity_model/`.
- [requirements.txt](requirements.txt) — runtime deps mirrored to
  [annuity_model/requirements.txt](annuity_model/requirements.txt). Test/dev
  tools live in [annuity_model/requirements-dev.txt](annuity_model/requirements-dev.txt).
- [DOCUMENTATION_MAP.md](DOCUMENTATION_MAP.md) — complete inventory of tracked
  docs with one-line purpose per file.

## Where to start

| You are a… | Read first |
|------------|------------|
| Human dev on macOS | [MACOS_HANDOFF.md](MACOS_HANDOFF.md) |
| Human dev on Windows | [AGENTS.md](AGENTS.md) (Windows section) |
| Human dev needing full governance map | [PROJECT_DEVELOPMENT_GUIDE.md](PROJECT_DEVELOPMENT_GUIDE.md) |
| AI agent / Cursor session | [AGENTS.md](AGENTS.md) → [annuity_model/AGENTS.md](annuity_model/AGENTS.md) → `annuity_model/.cursor/rules/*.mdc` |
| AI agent on any platform (Cursor/Claude Code/Codex) | [PROJECT_DEVELOPMENT_GUIDE.md](PROJECT_DEVELOPMENT_GUIDE.md) → [annuity_model/docs/AI_AGENT_PREFLIGHT.md](annuity_model/docs/AI_AGENT_PREFLIGHT.md) |
| Actuary / model owner | [annuity_model/docs/model_parity_contract.md](annuity_model/docs/model_parity_contract.md), [annuity_model/docs/rila_parity_contract.md](annuity_model/docs/rila_parity_contract.md), [annuity_model/docs/rila_product_spec.md](annuity_model/docs/rila_product_spec.md) |
| Release manager | [annuity_model/docs/parity_test_checklist.md](annuity_model/docs/parity_test_checklist.md) |

## Bootstrap

```bash
# macOS (Apple Silicon)
cd annuity_model && ./bootstrap_macos.sh

# Windows
cd annuity_model && bootstrap.bat
```

After bootstrap, the venv is at `annuity_model/.venv/`.

## Daily commands

```bash
cd annuity_model
pytest tests/ tests/parity/ -q          # full suite (order ~1 min locally; CI varies)
pytest tests/parity/ -v                 # parity gates only
python scripts/deep_smoke.py            # build + validate real .xlsx for every product
./run_pricing_ui.sh                     # launch Streamlit UI on :8501
./run_test_dashboard.sh                 # launch test dashboard on :8502
./run_tests_report.sh                   # generate reports/pytest_report.html
```

Windows equivalents: `bootstrap.bat`, `run_pricing_ui.bat`, `run_test_dashboard.bat`,
`run_tests_report.bat`.

## Invariants (do not break)

- **Every** `wb.save(...)` in an Excel builder MUST be preceded by
  `validate_workbook_or_raise(wb)`. See
  [annuity_model/excel_workbook_validator.py](annuity_model/excel_workbook_validator.py)
  and [annuity_model/.cursor/rules/excel-formula-safety.mdc](annuity_model/.cursor/rules/excel-formula-safety.mdc).
- Python ↔ Excel parity tolerances are codified in
  [annuity_model/docs/model_parity_contract.md](annuity_model/docs/model_parity_contract.md)
  (SPIA ALM) and
  [annuity_model/docs/rila_parity_contract.md](annuity_model/docs/rila_parity_contract.md).
  Cross-check before changing any tolerance literal.
- RILA's liability sheet uses column **M** for `ExpTotalCF` (everywhere else
  it is column **S**). The ALM linker takes a `liability_total_col` argument —
  do not hardcode column letters.
- Disinvestment tie-break: lower-indexed bucket wins. Excel threshold is
  `5e-10` (half the `1e-9` epsilon interval). See contract §2.

## License / data

The repo does **not** ship RP-2014 / MP-2016 mortality tables (proprietary to
SOA). Provide your own `q_x` CSV when running the engine; see the module
docstring at the top of [annuity_model/pricing_projection.py](annuity_model/pricing_projection.py).
