# Platform capability gap register

This register maps current platform maturity to a target state across
model governance, actuarial capability, and software architecture.
It is intended to be updated quarterly and referenced in release planning.

## Maturity scale

- **L1 - ad hoc:** inconsistent or undocumented.
- **L2 - foundational:** controlled but partially manual.
- **L3 - managed:** standardized and mostly automated.
- **L4 - optimized:** proactive, monitored, and continuously improved.

## Current assessment (updated 2026-05-10)

| Domain | Capability | Current | Target | Gap summary | Primary evidence |
|---|---|---:|---:|---|---|
| Governance | Independent challenge | L2 | L4 | CODEOWNERS routing exists, but enforceable second-party review is not yet active under solo-owner branch protection. | [`.github/CODEOWNERS`](../../.github/CODEOWNERS), [`.github/branch-protection.json`](../../.github/branch-protection.json) |
| Governance | Model inventory and tiering | L2 | L3 | Model inventory now exists, but maturity labels and limitation updates still require ongoing synchronization with implementation. | [`model_inventory.md`](model_inventory.md) |
| Governance | Assumption approval workflow | L2 | L4 | Technical artifact registry is strong, but approval/challenger/validity metadata is not yet formalized. | [`data_registry.py`](../src/annuity_model/data_registry.py) |
| Governance | Placeholder release controls | L3 | L4 | Release guardrail and waiver evidence exist; remaining gap is independent approval/expiry lifecycle automation. | [`scripts/check_assumption_release_guardrails.py`](../scripts/check_assumption_release_guardrails.py), [`.release/assumption_waiver.md`](../.release/assumption_waiver.md) |
| Governance | Ongoing performance monitoring | L2 | L4 | Strong pre-merge gates exist; post-merge KRIs and drift governance are not yet formalized. | [`actuarial_benchmarks.md`](actuarial_benchmarks.md) |
| Actuarial | Product coverage breadth | L3 | L4 | Ten products are implemented; out-of-scope features and simplifications remain by design in some products. | [`rila_product_spec.md`](rila_product_spec.md), [`tests/test_regression_matrix.py`](../tests/test_regression_matrix.py) |
| Actuarial | Portfolio heterogeneity fidelity | L2 | L4 | Shared scenario package can under-represent heterogeneous cohorts in mixed inforce books. | [`portfolio_runner_spec.md`](portfolio_runner_spec.md) |
| Actuarial | Dynamic policyholder behavior | L2 | L4 | Lapse framework is currently static v1 with explicit future v2 scope. | [`lapse_framework.md`](lapse_framework.md) |
| Actuarial | Experience studies/backtesting | L2 | L3 | An O/E calculation module exists; remaining gap is governed ingestion, runbook, and assumption-update workflow integration. | [`experience_study.py`](../src/annuity_model/experience_study.py), [`model_change_log.md`](model_change_log.md) |
| Software | Product extensibility wiring | L3 | L4 | `ProductDefinition` is the canonical source of truth and legacy views are derived; remaining gap is reducing compatibility surface area as APIs stabilize. | [`products/__init__.py`](../src/annuity_model/products/__init__.py), [`product_registry.py`](../src/annuity_model/product_registry.py), [`product_excel.py`](../src/annuity_model/product_excel.py), [`liability_dispatch.py`](../src/annuity_model/liability_dispatch.py) |
| Software | UI modularity and maintainability | L2 | L4 | App shell, overview, diagnostics, routing, and badges are extracted; pricing, ALM, what-if, and Excel page bodies remain in the large orchestration module. | [`pricing_ui.py`](../src/annuity_model/pricing_ui.py), [`ui/MIGRATION.md`](../src/annuity_model/ui/MIGRATION.md) |
| Software | Run reproducibility storage | L3 | L3 | UI/CLI pricing and portfolio runs persist SQLite ledger records and workbook evidence by default; continue hardening export/replay ergonomics. | [`run_ledger.py`](../src/annuity_model/run_ledger.py), [`workbook_run_evidence.py`](../src/annuity_model/workbook_run_evidence.py), [`cli.py`](../src/annuity_model/cli.py) |
| Software | Programmatic integration API | L1 | L3 | No service API for enterprise automation; UI/CLI only today. | [`cli.py`](../src/annuity_model/cli.py), [`pricing_ui.py`](../src/annuity_model/pricing_ui.py) |

## Prioritized remediation order

1. Governance controls that reduce model risk quickly (inventory, assumptions, guardrails, challenge).
2. Actuarial fidelity upgrades that materially improve portfolio realism.
3. Software modularity and durable run infrastructure to improve delivery velocity and auditability.
4. API and operational monitoring automation.

## Quarterly update checklist

- Re-score all capabilities against actual controls in code and CI.
- Record promoted controls and residual risk acceptances.
- Append major score changes to [`model_change_log.md`](model_change_log.md) when model-impacting.
