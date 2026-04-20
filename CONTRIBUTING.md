# Contributing to Code_Sandbox

Thanks for working on the actuarial models in this workspace. This file is the
short, human-facing version of the contribution flow. The detailed,
agent-facing version lives in [`AGENTS.md`](AGENTS.md) and
[`annuity_model/AGENTS.md`](annuity_model/AGENTS.md). For a complete map of
rules, skills, subagent/delegation behavior, regression gates, and control
framework, read [`PROJECT_DEVELOPMENT_GUIDE.md`](PROJECT_DEVELOPMENT_GUIDE.md).

## TL;DR before you open a PR

```bash
# From the repo root, on a branch off main:
cd annuity_model
python -m pytest tests/parity -q                      # 1. parity gate
python -m pytest -q                                   # 2. full unit tests
python scripts/deep_smoke.py                          # 3. end-to-end smoke
python scripts/render_parity_contract.py --check      # 4. tolerance docs in sync
pre-commit run --all-files                            # 5. lint / format / type
```

All five must exit 0. CI will rerun them on every push to your PR.

## Where to land changes

| Kind of change                   | Land in                                                       |
|----------------------------------|---------------------------------------------------------------|
| New product engine               | New `annuity_model/<product>_projection.py` (see README)      |
| New Excel builder                | New `annuity_model/build_<product>_excel_workbook.py`         |
| New parity test                  | `annuity_model/tests/parity/test_<product>_parity.py`         |
| Tolerance change                 | `annuity_model/parity_constants.py` + `model_change_log.md`   |
| New Streamlit page               | `annuity_model/ui/pages/` (see `ui/MIGRATION.md`)             |
| Validator / arity / func entry   | `annuity_model/excel_workbook_validator.py`                   |
| Documentation                    | `annuity_model/docs/` (READMEs, glossary, runbooks)           |

## Hard rules

1. **Never weaken a tolerance to make a test pass.** Tolerances live in
   `parity_constants.py` and route through CODEOWNERS + `model_change_log.md`.
   This is mechanically enforced in CI by
   [`annuity_model/scripts/check_parity_constants_changelog.py`](annuity_model/scripts/check_parity_constants_changelog.py)
   in the `pre-commit (lint + format + mypy)` job — any PR that modifies
   `parity_constants.py` without also touching `docs/model_change_log.md`
   fails the gate.
2. **Every `wb.save(...)` is preceded by `validate_workbook_or_raise(wb)`.**
   The static validator is the only thing protecting users from malformed
   Excel formulas.
3. **Excel column letters come from `liability_layouts.py`.** Hardcoding `"S"`
   or `"M"` outside that module is the bug class the registry exists to
   prevent.
4. **No raw `t_rem` ordering.** Always use the epsilon-adjusted argsort key
   documented in `docs/model_parity_contract.md` section 2.
5. **Every numerical bug becomes a permanent regression test** under
   `tests/parity/` with a clear `# Regression: YYYY-MM-DD` comment.
6. **Cross-platform launchers are added in pairs.** Every new `*.sh` ships
   alongside its `*.bat` twin in the same commit.

## Useful runbooks

* Validator failed: [`docs/runbooks/debug_validator_failure.md`](annuity_model/docs/runbooks/debug_validator_failure.md)
* Parity failed: [`docs/runbooks/investigate_parity_break.md`](annuity_model/docs/runbooks/investigate_parity_break.md)
* Excel cache problem: [`docs/runbooks/regenerate_excel_cache.md`](annuity_model/docs/runbooks/regenerate_excel_cache.md)
* Cutting a release: [`docs/runbooks/release.md`](annuity_model/docs/runbooks/release.md)

## CODEOWNERS

The `.github/CODEOWNERS` file gates changes to the parity-critical surface,
the Excel builders, the parity contracts and tests, and CI / agent guidance.
The rationale for each block is documented in
[`annuity_model/docs/CODEOWNERS_RATIONALE.md`](annuity_model/docs/CODEOWNERS_RATIONALE.md).
If your PR touches a CODEOWNERS-gated path, expect the owner-review
requirement and don't try to route around it.

## Local setup (cross-platform)

The full bootstrap procedures live in:
* macOS / Apple Silicon: `annuity_model/bootstrap_macos.sh`
* Windows / PowerShell: `MACOS_HANDOFF.md` (Windows section at the bottom)

Both install the pinned runtime + dev requirements; nothing else is needed.

## Questions

Read [`annuity_model/docs/glossary.md`](annuity_model/docs/glossary.md) first.
For everything else, open a draft PR with the question in the description.
