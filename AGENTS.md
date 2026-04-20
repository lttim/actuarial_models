# Agent Instructions — Actuarial Models Workspace

This workspace contains actuarial pricing and ALM projection models. Each product lives in
its own subdirectory and has its own parity contract between the Python engine and generated
Excel workbooks.

For a consolidated development-governance map that is explicit for both humans and
cross-platform AI agents (Cursor, Claude Code, Codex, etc.), read
`PROJECT_DEVELOPMENT_GUIDE.md`.

## Workspace structure

| Directory | Product |
|-----------|---------|
| `annuity_model/` | Multi-product pricing + ALM (SPIA, Term, RILA, MYGA, FIA, VA, WL, UL, IUL, VUL); see `annuity_model/README.md` |
| `actuarial_parity_kit/` | Reusable governance template for new products |

## Universal rules for all actuarial products in this workspace

1. **Python = Excel**: Every product must maintain exact numerical parity between its Python
   calculation engine and the Excel workbook it generates. See each product's
   `docs/model_parity_contract.md`.
2. **Test gates**: the canonical "before completing any task" gates live in
   [annuity_model/AGENTS.md -- "Before completing any task"](annuity_model/AGENTS.md#before-completing-any-task----canonical-gates).
   This file does not duplicate the commands; follow the link.
3. **Epsilon tie-breaking**: Never use raw floating-point ordering when values are nominally
   equal. Always use index-based epsilon offsets (see parity contract for specification).
4. **Step-level validation**: Validate month-by-month intermediate state, not only final output.
5. **New bug = new test**: Every numerical bug must produce a permanent regression test.
6. **Reuse the kit**: When starting a new product, copy `actuarial_parity_kit/` into the
   new repo and adapt. Do not start from scratch.
7. **Tolerance constants live in code, not docs**: `annuity_model/parity_constants.py` is the
   single source of truth. The two parity contracts and the release checklist render their
   tolerance tables from this module via `scripts/render_parity_contract.py`. Never edit a
   rendered table by hand.

## Starting a new product repo

```
cp -r actuarial_parity_kit/ ../new_product_model/
cd ../new_product_model/
# Rename and adapt: cursor_rules/ → .cursor/rules/, then fill in product-specific logic
```

## Cross-platform setup (Windows + macOS Apple Silicon)

This repo is a single Git repository at `Code_Sandbox/`. There must **not** be a
nested `.git` inside `annuity_model/` or any other product directory.

Line endings, ignore rules, and binary handling are pinned in the root
`.gitattributes` and `.gitignore`. Do not override `core.autocrlf` per machine —
the attributes file already does the right thing for both OSes.

### Cloning on macOS (Apple Silicon, M-series)

```bash
git clone https://github.com/lttim/actuarial_models.git Code_Sandbox
cd Code_Sandbox/annuity_model
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
./run_pricing_ui.sh        # Streamlit UI
pytest tests/ tests/parity/ # full regression
```

All Python deps (numpy, pandas, openpyxl, pyarrow, streamlit, matplotlib) ship
arm64 wheels for Python 3.11+, so the M-series Mac uses the native build with
no Rosetta. If `pip install` ever falls back to building from source, upgrade
`pip` first (`python3 -m pip install --upgrade pip`).

### Cloning on Windows

```powershell
git clone https://github.com/lttim/actuarial_models.git Code_Sandbox
cd Code_Sandbox\annuity_model
py -3 -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
.\run_pricing_ui.bat
py -3 -m pytest tests\ tests\parity\
```

### Launcher parity

Every shell launcher must have a Windows twin and vice-versa. Current pairs:

| Purpose             | Windows                  | macOS / Linux            |
|---------------------|--------------------------|--------------------------|
| Pricing Streamlit   | `run_pricing_ui.bat`     | `run_pricing_ui.sh`      |
| Test dashboard      | `run_test_dashboard.bat` | `run_test_dashboard.sh`  |
| pytest HTML report  | `run_tests_report.bat`   | `run_tests_report.sh`    |

When adding a new launcher, ship both files in the same commit and keep them in
sync. `.gitattributes` enforces `*.sh = LF` and `*.bat = CRLF` so neither side
gets mangled by the other OS's checkout settings.
