# Agent Instructions — Actuarial Models Workspace

This workspace contains actuarial pricing and ALM projection models. Each product lives in
its own subdirectory and has its own parity contract between the Python engine and generated
Excel workbooks.

## Workspace structure

| Directory | Product |
|-----------|---------|
| `annuity_model/` | SPIA (Single Premium Immediate Annuity) |
| `actuarial_parity_kit/` | Reusable governance template for new products |

## Universal rules for all actuarial products in this workspace

1. **Python = Excel**: Every product must maintain exact numerical parity between its Python
   calculation engine and the Excel workbook it generates. See each product's
   `docs/model_parity_contract.md`.
2. **Test gates**: `pytest tests/parity/` must pass at 0.00 discrepancy before any merge.
3. **Epsilon tie-breaking**: Never use raw floating-point ordering when values are nominally
   equal. Always use index-based epsilon offsets (see parity contract for specification).
4. **Step-level validation**: Validate month-by-month intermediate state, not only final output.
5. **New bug = new test**: Every numerical bug must produce a permanent regression test.
6. **Reuse the kit**: When starting a new product, copy `actuarial_parity_kit/` into the
   new repo and adapt. Do not start from scratch.

## Starting a new product repo

```
cp -r actuarial_parity_kit/ ../new_product_model/
cd ../new_product_model/
# Rename and adapt: cursor_rules/ → .cursor/rules/, then fill in product-specific logic
```
