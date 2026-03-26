# Actuarial Parity Kit

A reusable governance template for actuarial models that maintain numerical parity
between a Python calculation engine and a generated Excel workbook.

## What's included

| Path | Purpose |
|------|---------|
| `cursor_rules/actuarial-parity.mdc` | Cursor AI rule — copy to `.cursor/rules/` |
| `docs/model_parity_contract_template.md` | Canonical contract template |
| `docs/parity_test_checklist_template.md` | Release gate checklist template |
| `AGENTS_template.md` | AGENTS.md template for the product repo |
| `tests/parity/excel_formula_sim_template.py` | Excel formula simulation base class |
| `tests/parity/test_alm_parity_template.py` | Parity test template |
| `scripts/export_parity_trace.py` | Month-by-month state dump for debugging |

## How to use in a new product repo

```bash
# From your new product repo root:
cp -r ../actuarial_parity_kit/cursor_rules/ .cursor/rules/
cp ../actuarial_parity_kit/AGENTS_template.md AGENTS.md
mkdir -p docs tests/parity
cp ../actuarial_parity_kit/docs/*_template.md docs/
cp ../actuarial_parity_kit/tests/parity/* tests/parity/
cp ../actuarial_parity_kit/scripts/* scripts/
```

Then:
1. Rename `*_template.md` → `*.md` and fill in product-specific content.
2. Implement `excel_formula_sim.py` to simulate your product's Excel formulas.
3. Adapt `test_alm_parity.py` with your product's golden scenarios.
4. Run `pytest tests/parity/ -v` — all should pass before first merge.

## Core principle

Every actuarial model that has both a Python engine and an Excel output must:
1. Define explicit tolerances for every tracked variable.
2. Specify a deterministic ordering policy for any tie-break situation.
3. Document the epsilon strategy for floating-point edge cases.
4. Test step-by-step (monthly), not only the final output.
5. Convert every discovered bug into a permanent regression test.

## The most common failure modes (from production experience)

| Failure | Symptom | Root cause |
|---------|---------|------------|
| Double-sell | Excel sells 2× the needed amount | Epsilon threshold ≥ epsilon interval |
| Wrong bucket | Python sells higher-indexed bucket | Raw float comparison of accumulated t_rem |
| Formula double-DF | Excel face value treated as MV | Missing or extra division by DF |
| Stale cache | data_only=True returns old values | openpyxl reads pre-recalculation cache |
| Hidden offset | Final output matches, intermediates don't | Compensating errors in different months |
