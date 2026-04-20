# Platform engineering roadmap

This roadmap defines implementation tracks for software architecture
hardening and delivery scalability.

## Track 1: UI modularization

### Goal

Reduce regression coupling by splitting the Streamlit monolith into
feature modules.

### Proposed decomposition

- `ui/run_page.py`
- `ui/portfolio_page.py`
- `ui/alm_page.py`
- `ui/what_if_page.py`
- `ui/excel_page.py`
- `ui/diagnostics_page.py`

### Near-term tasks

- Extract pure rendering/helpers from `pricing_ui.py`.
- Maintain `pricing_run_form_state.py` as the key contract layer.
- Preserve AppTest E2E coverage and add page-level unit tests.

## Track 2: canonical product-definition pipeline

### Goal

Consolidate multiple product wiring registries into one canonical definition source.

### Proposed direction

- Keep `products/<name>/__init__.py` as canonical source.
- Generate/derive wiring for:
  - pricing adapters
  - workbook builders
  - liability-path conversion
  - formatter metadata
- Remove synchronization burden currently enforced by meta-invariant tests.

### Near-term tasks

- Introduce a `ProductDefinition` schema with all required hooks.
- Add migration compatibility layer to avoid breaking current imports.

## Track 3: durable run/artifact store

### Goal

Make runs reproducible and audit-ready without relying on ad hoc files.

### Proposed MVP

- File-backed or SQLite run ledger with:
  - run_id
  - timestamp
  - model version
  - scenario identifiers
  - key inputs hash
  - output artifact paths
  - parity/gate status

### Integration points

- [`cli.py`](../cli.py): persist run records after `portfolio-run`.
- [`pricing_ui.py`](../pricing_ui.py): optional persistence toggle for user runs.

## Track 4: API layer for orchestration

### Goal

Expose programmatic entry points for enterprise/batch integration.

### Proposed MVP endpoints

- `POST /pricing/run`
- `POST /portfolio/run`
- `GET /runs/{run_id}`
- `GET /health`

### Constraints

- Keep Python engine as source of truth.
- Reuse existing validation and scenario materialization.
- Return run IDs linked to durable ledger entries.

## Delivery guardrails

- No parity contract weakening during refactors.
- New abstractions require regression and invariant tests.
- Preserve existing CLI/UI behavior while migrating internals.
