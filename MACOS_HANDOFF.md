# macOS Handoff — Cursor on Apple Silicon

You're on a fresh Cursor + macOS install (MacBook Air, M5). Everything you
need to continue from where the Windows session left off is already in this
repository. **Read this file first**, then follow the bootstrap steps below.

The single source of truth is the GitHub repo
[`lttim/actuarial_models`](https://github.com/lttim/actuarial_models). There
is no separate Cursor profile to migrate, no out-of-band rules to install:
the Cursor rules, AGENTS.md instructions, and parity / Excel-validator
controls are all checked in.

---

## 1. Clone

```bash
mkdir -p ~/Code
cd ~/Code
git clone https://github.com/lttim/actuarial_models.git Code_Sandbox
cd Code_Sandbox
```

> **Open `Code_Sandbox/` (the repo root, not `annuity_model/`) as the Cursor
> workspace.** That's the level where `AGENTS.md` and both product
> directories live, so Cursor auto-discovers all rules.

## 2. Bootstrap the Python environment

```bash
cd annuity_model
./bootstrap_macos.sh
```

This script will:

- check for Python 3.11+ and recommend `brew install python@3.12` if missing,
- create `annuity_model/.venv` (gitignored),
- install `requirements.txt` into it,
- run the full regression suite (`pytest tests/ tests/parity/`) — **must exit 0**
  (summary line lists `passed` and any `skipped`; the pass count grows with the suite),
- print the next-step commands.

If you would rather do it by hand:

```bash
cd annuity_model
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
pytest tests/ tests/parity/
```

All Python deps (`numpy`, `pandas`, `pyarrow`, `openpyxl`, `streamlit`,
`matplotlib`) ship native arm64 wheels for Python 3.11+, so the M-series
Mac runs them without Rosetta.

## 3. Verify (smoke test)

```bash
cd annuity_model
source .venv/bin/activate
./run_pricing_ui.sh                  # opens Streamlit on http://localhost:8501
./run_tests_report.sh                # writes reports/pytest_report.html and opens it
```

If the Streamlit UI loads and the HTML report shows the full suite green (no
failures), you're aligned with CI expectations.

---

## 4. What Cursor needs to know (and where it reads it from)

Open `Code_Sandbox/` in Cursor. Rules and skills load automatically from these
checked-in files — you do **not** need to copy anything from `~/.cursor/`:

| File / dir | Scope | What it enforces |
|---|---|---|
| `AGENTS.md` (repo root) | Whole workspace | Workspace structure, parity rule #1, cross-platform setup, launcher pairs. |
| `annuity_model/AGENTS.md` | Product | Parity gates, key files, **mandatory `excel_workbook_validator.validate_workbook_or_raise(wb)`** before every `wb.save(...)`, cross-sheet column rules, `liability_total_col` guidance. |
| `annuity_model/.cursor/rules/actuarial-parity.mdc` | Always-on | Tie-break / epsilon / step-level parity invariants. |
| `annuity_model/.cursor/rules/excel-formula-safety.mdc` | Globbed to builders | Off-by-one parens, missing `IF` false branch, trailing empty args, wrong `Liabilities!` column letter. |
| `.cursor/rules/handoff-recall.mdc` (repo root) | Workspace | Canonical `!handoff` / `!recall` — writes **gitignored** files under `.cursor/handoffs/`. |
| `annuity_model/.cursorrules` | Product | Legacy `!handoff` hook — overwrites `annuity_model/state.md` when used; prefer root handoff protocol when both exist. |
| `annuity_model/state.md` | Session snapshot | Short human handoff; optional if you use `.cursor/handoffs/`. |
| `PROJECT_DEVELOPMENT_GUIDE.md` | Workspace | Governance map for humans and non-Cursor agents. |
| `annuity_model/docs/model_parity_contract.md` | Reference | SPIA/ALM parity tolerances, tie-break, epsilon policy. |
| `annuity_model/docs/rila_parity_contract.md` | Reference | RILA Python ↔ Excel parity addendum. |
| `annuity_model/docs/rila_product_spec.md` | Reference | RILA v1 product definition. |
| `.gitattributes` (repo root) | Git | LF for sources, CRLF for `.bat`, `+x` for `.sh`, binary for `.xlsx`/`.png`/etc. |
| `.gitignore` (repo root) | Git | `.venv`, `__pycache__`, `model output/`, `reports/`, macOS `.DS_Store`/`._*`, Windows `Thumbs.db`. |

The Cursor "skills" (canvas, babysit, create-rule, etc.) and built-in MCP
servers (`cursor-app-control`, `cursor-ide-browser`) ship with Cursor itself,
so a fresh macOS install already has them. Nothing to migrate.

## 5. First prompt for the new Cursor session

Open the chat in `Code_Sandbox/` and paste:

> Read `MACOS_HANDOFF.md`, `annuity_model/AGENTS.md`,
> `annuity_model/state.md`, and the two rule files under
> `annuity_model/.cursor/rules/`. Then run
> `cd annuity_model && ./bootstrap_macos.sh` and report the pytest summary
> line (`passed` / `skipped`, exit code).

That gives the agent the full picture: invariants, current state, and a
verified-green baseline to start from.

---

## 6. Project invariants (cheat sheet)

These are restated here so you can scan them in one place. The full versions
live in the files referenced in §4.

1. **Single source of truth.** There is exactly one `.git` directory
   (`Code_Sandbox/.git`). Never re-init a nested repo inside `annuity_model/`
   or any other product directory.
2. **Python = Excel.** Every code change must keep the Python engine and the
   generated Excel workbook at exact parity. `pytest tests/parity/ -v` must be
   0.00 discrepancy before any merge.
3. **Excel formulas pass static validation before save.** Every workbook
   builder calls `excel_workbook_validator.validate_workbook_or_raise(wb)`
   immediately before `wb.save(...)`. The validator now checks syntax and
   cross-sheet semantics (incl. references hidden inside `INDIRECT(...)`).
   When wiring a new product into the shared ALM helpers, pass the right
   `liability_total_col` — RILA uses `"M"`, SPIA / Term Life use `"S"`.
4. **Launcher pairs.** Every Windows `.bat` ships with a matching macOS/Linux
   `.sh` in the same commit. `.gitattributes` enforces line endings so neither
   side gets corrupted on checkout.
5. **New bug = new test.** Every numerical or rendering bug must produce a
   permanent regression test.

## 7. Where the most recent work stopped

- **`git log -1 --oneline`** at the repo root shows the current tip of `main`.
- **`annuity_model/state.md`** is a short optional snapshot (may lag `main`).
- **`.cursor/handoffs/*.md`** holds structured cross-chat handoffs from `!handoff`
  (gitignored; not in the clone unless you create them locally).

## 8. Common gotchas on first macOS run

- **`./bootstrap_macos.sh: Permission denied`** — the file is checked in with
  `100755` mode, so this should not happen. If it does, run
  `chmod +x bootstrap_macos.sh run_*.sh` once and re-run.
- **`zsh: command not found: python`** — macOS only ships `python3`. The
  launchers prefer `python3`; you only need plain `python` if you're in a
  venv that creates the `python` symlink (`source .venv/bin/activate` does
  that).
- **`xcrun: error: invalid active developer path`** when `pip install` builds
  from source — install the Command Line Tools (`xcode-select --install`).
  This is rare; arm64 wheels exist for everything in `requirements.txt`.
- **Cursor doesn't see the rules** — make sure you opened `Code_Sandbox/` as
  the workspace root, not `annuity_model/` or the parent of `Code_Sandbox/`.
- **Wrong Python for tooling** — run gates with `annuity_model/.venv/bin/python3`
  after bootstrap; bare macOS `python3` may be 3.9 and lack modern typing
  features the repo requires.

## 9. Pushing back to GitHub

Set git identity once on the new machine if you haven't:

```bash
git config --global user.name  "Your Name"
git config --global user.email "you@example.com"
```

The remote `origin` is already set in the repo (`git remote -v` will show
`https://github.com/lttim/actuarial_models.git`). For pushes you'll need to
authenticate with HTTPS + a Personal Access Token, or switch to SSH:

```bash
git remote set-url origin git@github.com:lttim/actuarial_models.git
```
