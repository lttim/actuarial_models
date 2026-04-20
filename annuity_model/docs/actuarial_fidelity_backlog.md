# Actuarial fidelity enhancement backlog

This backlog translates the platform assessment into implementable
actuarial capability improvements.

## Priority A: cohort-aware portfolio scenarios

### Objective

Improve mixed-book realism by supporting cohort-level assumption packages
within a single portfolio run.

### Scope

- Add cohort keys (sex, smoker, issue-year band, product family) in inforce parsing.
- Materialize per-cohort scenario bundles using existing scenario builders.
- Aggregate results while preserving current rollup invariants.

### Candidate implementation surfaces

- [`inforce_parsers.py`](../inforce_parsers.py)
- [`pricing_scenario_materialize.py`](../pricing_scenario_materialize.py)
- [`portfolio_runner.py`](../portfolio_runner.py)
- [`docs/portfolio_runner_spec.md`](portfolio_runner_spec.md)

### Evidence artifacts

- New parity tests in `tests/parity/portfolio`.
- Benchmark deltas recorded in [`docs/actuarial_benchmarks.md`](actuarial_benchmarks.md).

## Priority B: dynamic lapse framework v2

### Objective

Move from static lapse assumptions to behaviorally responsive lapse models.

### Scope

- Add optional dynamic lapse strategy interface.
- Support interest-rate and in-the-moneyness sensitivity.
- Define surrender-charge interaction and recapture treatment.

### Candidate implementation surfaces

- [`lapse.py`](../lapse.py)
- [`docs/lapse_framework.md`](lapse_framework.md)
- Product engines (`*_projection.py`) that support lapse assumptions.

### Evidence artifacts

- Property/invariant tests for lapse monotonicity and bounds.
- Regression matrix entries for dynamic lapse scenarios.

## Priority C: scenario governance and stress catalog

### Objective

Establish governed scenario sets with reproducibility and clear intended use.

### Scope

- Add scenario catalog metadata (owner, seed, effective dates, purpose).
- Distinguish pricing base, stress, and capital-style scenarios.
- Provide deterministic replay identifiers in outputs.

### Candidate implementation surfaces

- [`data_registry.py`](../data_registry.py)
- [`pricing_scenario_materialize.py`](../pricing_scenario_materialize.py)
- CLI/export payloads in [`cli.py`](../cli.py)

### Evidence artifacts

- Scenario catalog document and checks.
- Audit sample showing run replay from stored scenario identifiers.

## Priority D: experience study and backtesting loop

### Objective

Create a managed observed-vs-expected loop for assumption updates.

### Scope

- Define observed data schema and ingestion format.
- Calculate O/E metrics by product/cohort/time bucket.
- Set threshold-based triggers for assumption review.

### Candidate implementation surfaces

- New module: `experience_study.py` (planned).
- New runbook for O/E operations.
- Link updates into [`docs/model_change_log.md`](model_change_log.md).

### Evidence artifacts

- Reproducible O/E report sample with signoff.
- Logged assumption decisions referencing O/E outcomes.
