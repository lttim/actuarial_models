# Justfile -- one-line recipes for everyday work.
# Install just: https://github.com/casey/just
# List recipes: `just --list` (or `just`).

set shell := ["bash", "-uc"]
set positional-arguments

# Default: list available recipes.
default:
    @just --list

# Bootstrap the macOS dev environment (Homebrew, venv, deps, smoke check).
bootstrap:
    bash annuity_model/bootstrap_macos.sh

# Install / upgrade dependencies into the active venv.
install:
    pip install --upgrade pip
    pip install -r annuity_model/requirements.lock -r annuity_model/requirements-dev.txt

# Run the parity-critical subset (blocks any merge on failure).
parity:
    cd annuity_model && python -m pytest tests/parity -q

# Run the full unit-test suite.
test *args="":
    cd annuity_model && python -m pytest -q "$@"

# End-to-end smoke (3 products + full Excel validator).
smoke:
    cd annuity_model && python scripts/deep_smoke.py

# Verify tolerance docs are in sync with parity_constants.py.
docs-check:
    cd annuity_model && python scripts/render_parity_contract.py --check

# Regenerate tolerance docs from parity_constants.py.
docs-render:
    cd annuity_model && python scripts/render_parity_contract.py

# Build the mkdocs site under annuity_model/site/.
docs-build:
    cd annuity_model && mkdocs build --strict

# Serve the mkdocs site locally on http://127.0.0.1:8000.
docs-serve:
    cd annuity_model && mkdocs serve

# Launch the Streamlit pricing UI.
ui:
    cd annuity_model && bash run_pricing_ui.sh

# Run pre-commit on every file.
lint:
    pre-commit run --all-files

# Performance benchmarks (advisory; the parity gate does not depend on these).
bench:
    cd annuity_model && python -m pytest tests/test_perf_baselines.py --benchmark-only

# CI-equivalent local run: lint + tests + smoke + docs check.
ci: lint test smoke docs-check
    @echo "All CI gates passed locally."

# AI-agent / contributor pre-merge gate: run the four canonical gates from
# annuity_model/AGENTS.md in order and print "READY TO COMMIT" only when all
# four exit 0. The PR template (.github/pull_request_template.md) lists the
# same four gates as checkboxes; this recipe is the one-liner that ticks
# them all in a single command.
preflight:
    @echo "[1/4] parity gate"
    @cd annuity_model && python -m pytest tests/parity -q
    @echo "[2/4] full unit-test suite"
    @cd annuity_model && python -m pytest -q
    @echo "[3/4] end-to-end smoke (10 products + Excel validator)"
    @cd annuity_model && python scripts/deep_smoke.py
    @echo "[4/4] tolerance + actuarial-benchmark docs in sync"
    @cd annuity_model && python scripts/render_parity_contract.py --check
    @cd annuity_model && python scripts/render_actuarial_benchmarks.py --check
    @echo ""
    @echo "READY TO COMMIT: all four canonical gates passed."

# Container build + smoke (requires Docker).
docker-smoke:
    docker build -t annuity-model:dev .
    docker run --rm annuity-model:dev

# Trace a python-vs-excel discrepancy for the SPIA scenario.
trace steps="60":
    cd annuity_model && python scripts/parity_trace.py --steps {{steps}} --output traces/parity_trace.csv

# Refresh the lockfile from requirements.txt (use sparingly; CI should be the
# canonical source of truth).
lockfile:
    cd annuity_model && pip-compile --resolver=backtracking --output-file=requirements.lock requirements.txt

# Security: pip-audit + bandit on the package.
security:
    pip-audit --requirement annuity_model/requirements.lock --strict
    bandit -r annuity_model -x annuity_model/tests,annuity_model/.venv
