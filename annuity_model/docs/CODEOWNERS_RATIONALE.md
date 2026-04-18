# CODEOWNERS rationale

`/.github/CODEOWNERS` enumerates the files that require an owner review
before merge. This document explains *why* each block exists, so future
maintainers don't accidentally weaken a load-bearing gate.

## Parity-critical actuarial surface

| Path                                          | Why owner review is required                                              |
|-----------------------------------------------|---------------------------------------------------------------------------|
| `annuity_model/pricing_projection.py`         | Source of truth for SPIA cashflows + ALM. A silent change here is a release-stopping defect. |
| `annuity_model/term_projection.py`            | Term Life liability cashflow definition; consumed by Excel builder.       |
| `annuity_model/rila_projection.py`            | RILA crediting / cap / floor logic; consumed by Excel builder.            |
| `annuity_model/alm_excel_ladder.py`           | Generates the Excel formulas that must match Python -- any divergence breaks parity. |
| `annuity_model/excel_workbook_validator.py`   | The static gate that catches malformed formulas. Loosening it loses safety. |
| `annuity_model/product_registry.py`           | The dispatch surface that lets a new product land in 2 files; structural correctness matters here more than perf. |
| `annuity_model/parity_constants.py`           | The immutable tolerance contract. Every value here is referenced by code, tests, and rendered docs. |

## Excel builders

| Path                                          | Why                                                                       |
|-----------------------------------------------|---------------------------------------------------------------------------|
| `build_pricing_excel_workbook.py` (SPIA)      | One off-by-one in the liability column letter and ModelCheck explodes.    |
| `build_term_excel_workbook.py`                | Same.                                                                     |
| `build_rila_excel_workbook.py`                | Same; plus RILA must use column M, not S.                                 |
| `recalc_excel_shared.py`                      | Recalc helper used by parity tests; bug here masks formula errors.        |
| `product_excel.py`                            | Dispatcher across the three builders.                                     |

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
