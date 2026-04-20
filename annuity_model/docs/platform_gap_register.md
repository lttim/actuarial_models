# Platform capability gap register

This register maps current platform maturity to a target state across
model governance, actuarial capability, and software architecture.
It is intended to be updated quarterly and referenced in release planning.

## Maturity scale

- **L1 - ad hoc:** inconsistent or undocumented.
- **L2 - foundational:** controlled but partially manual.
- **L3 - managed:** standardized and mostly automated.
- **L4 - optimized:** proactive, monitored, and continuously improved.

## Baseline assessment (2026-04-19)

| Domain | Capability | Current | Target | Gap summary | Primary evidence |
|---|---|---:|---:|---|---|
| Governance | Independent challenge | L2 | L4 | CODEOWNERS routing exists, but enforceable second-party review is not yet active under solo-owner branch protection. | [`.github/CODEOWNERS`](../../.github/CODEOWNERS), [`.github/branch-protection.json`](../../.github/branch-protection.json) |
| Governance | Model inventory and tiering | L1 | L3 | No formal model inventory with tier/materiality, ownership, and review cadence metadata. | [`product_registry.py`](../product_registry.py) |
| Governance | Assumption approval workflow | L2 | L4 | Technical artifact registry is strong, but approval/challenger/validity metadata is not yet formalized. | [`data_registry.py`](../data_registry.py) |
| Governance | Placeholder release controls | L1 | L4 | Placeholder assumptions are documented but no hard release-time guardrail is mandated. | [`mortality_2017_cso.py`](../mortality_2017_cso.py), [`pricing_projection.py`](../pricing_projection.py) |
| Governance | Ongoing performance monitoring | L2 | L4 | Strong pre-merge gates exist; post-merge KRIs and drift governance are not yet formalized. | [`docs/actuarial_benchmarks.md`](actuarial_benchmarks.md) |
| Actuarial | Product coverage breadth | L3 | L4 | Ten products are implemented; out-of-scope features and simplifications remain by design in some products. | [`docs/rila_product_spec.md`](rila_product_spec.md), [`tests/test_regression_matrix.py`](../tests/test_regression_matrix.py) |
| Actuarial | Portfolio heterogeneity fidelity | L2 | L4 | Shared scenario package can under-represent heterogeneous cohorts in mixed inforce books. | [`docs/portfolio_runner_spec.md`](portfolio_runner_spec.md) |
| Actuarial | Dynamic policyholder behavior | L2 | L4 | Lapse framework is currently static v1 with explicit future v2 scope. | [`docs/lapse_framework.md`](lapse_framework.md) |
| Actuarial | Experience studies/backtesting | L1 | L3 | No formalized observed-vs-expected loop to govern assumption updates. | [`docs/model_change_log.md`](model_change_log.md) |
| Software | Product extensibility wiring | L3 | L4 | Registry framework is robust but still distributed across multiple legacy wiring surfaces. | [`product_registry.py`](../product_registry.py), [`product_excel.py`](../product_excel.py), [`liability_dispatch.py`](../liability_dispatch.py) |
| Software | UI modularity and maintainability | L2 | L4 | Main Streamlit surface remains monolithic and high-coupling. | [`pricing_ui.py`](../pricing_ui.py) |
| Software | Run reproducibility storage | L1 | L3 | Runs emit files/session outputs, but no durable run ledger for replay and audit. | [`cli.py`](../cli.py) |
| Software | Programmatic integration API | L1 | L3 | No service API for enterprise automation; UI/CLI only today. | [`cli.py`](../cli.py), [`pricing_ui.py`](../pricing_ui.py) |

## Prioritized remediation order

1. Governance controls that reduce model risk quickly (inventory, assumptions, guardrails, challenge).
2. Actuarial fidelity upgrades that materially improve portfolio realism.
3. Software modularity and durable run infrastructure to improve delivery velocity and auditability.
4. API and operational monitoring automation.

## Quarterly update checklist

- Re-score all capabilities against actual controls in code and CI.
- Record promoted controls and residual risk acceptances.
- Append major score changes to [`docs/model_change_log.md`](model_change_log.md) when model-impacting.
