# annuity_model documentation

Production-quality SPIA / Term Life / RILA pricing engine, ALM ladder, and Excel
workbook generator. Python is the source of truth; Excel is the auditor.

## Where to start

* **New to the codebase?** Read [Glossary](glossary.md) first, then the
  [SPIA / ALM parity contract](model_parity_contract.md).
* **Need the development governance map (humans + AI platforms)?** Read
  [Project development guide](../../PROJECT_DEVELOPMENT_GUIDE.md).
* **Adding a product?** The 2-file walkthrough lives in the
  [repository README](https://github.com/lttim/actuarial_models/blob/main/annuity_model/README.md).
* **Debugging a parity break?** Open
  [investigate_parity_break](runbooks/investigate_parity_break.md).
* **Validator failed?** Open [debug_validator_failure](runbooks/debug_validator_failure.md).
* **Cutting a release?** Open [release](runbooks/release.md).

## Architecture

```mermaid
flowchart LR
    SPIA[spia engine] --> Adapter
    Term[term engine] --> Adapter
    RILA[rila engine] --> Adapter
    Adapter --> Reg[ProductRegistry]
    Reg --> Builders[Excel builders]
    Builders --> Validator[excel_workbook_validator]
    Builders --> ALM[alm core]
    Reg --> UI[Streamlit ui/pages]
```

## Hard invariants

CI enforces all of these:

1. `parity_constants.MODELCHECK_TOL == 0.0` -- never weaken.
2. Every `wb.save(...)` call site is preceded by `validate_workbook_or_raise(wb)`.
3. RILA liability column is `M`; SPIA / Term liability column is `S`. Source
   of truth: `LIABILITY_LAYOUTS` in `liability_layouts.py`.
4. Tolerance tables in [model_parity_contract.md](model_parity_contract.md)
   and [rila_parity_contract.md](rila_parity_contract.md) are **generated**
   from `parity_constants.py`. Edit the constants, then run
   `python -m annuity_model.scripts.render_parity_contract`.

## Index

* [Glossary](glossary.md)
* Parity contracts: [SPIA / ALM](model_parity_contract.md), [RILA](rila_parity_contract.md)
* [RILA product spec](rila_product_spec.md)
* [Release checklist](parity_test_checklist.md)
* Runbooks: [parity break](runbooks/investigate_parity_break.md),
  [validator](runbooks/debug_validator_failure.md),
  [Excel cache](runbooks/regenerate_excel_cache.md),
  [release](runbooks/release.md),
  [assumption guardrail](runbooks/assumption_release_guardrail.md)
* Governance: [CHANGELOG](CHANGELOG.md), [model change log](model_change_log.md),
  [CODEOWNERS rationale](CODEOWNERS_RATIONALE.md),
  [model inventory](model_inventory.md),
  [assumption governance](assumption_governance.md),
  [gap register](platform_gap_register.md),
  [independent challenge activation](independent_challenge_activation.md)
* Roadmaps: [actuarial fidelity backlog](actuarial_fidelity_backlog.md),
  [platform engineering roadmap](platform_engineering_roadmap.md),
  [implementation phasing](implementation_phasing.md)
