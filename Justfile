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

# End-to-end smoke (all products in deep_smoke + full Excel validator).
smoke:
    cd annuity_model && python scripts/deep_smoke.py

# Verify tolerance docs are in sync with parity_constants.py.
docs-check:
    cd annuity_model && python scripts/render_parity_contract.py --check
    cd annuity_model && python scripts/check_documentation_map.py

# Release guardrail for placeholder/synthetic assumptions.
assumption-guardrail:
    cd annuity_model && python scripts/check_assumption_release_guardrails.py

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

# Post-stall confidence sweep for Excel generation, product parity, docs, and CI lint.
deep-assessment *args="":
    cd annuity_model && python scripts/deep_assessment.py "$@"

# AI-agent / contributor pre-merge gate: run the four canonical gates from
# annuity_model/AGENTS.md in order and print "READY TO COMMIT" only when all
# four exit 0. The PR template (.github/pull_request_template.md) lists the
# same four gates as checkboxes; this recipe is the one-liner that ticks
# them all in a single command.
#
# Note: gate 5 (Actuary SME review) is not invoked here -- it is a
# recursive gate that runs an autonomous fix-and-rereview loop and is
# triggered from the AI agent (via `!actuaryreview` or any
# natural-language "actuary review" request) per
# .cursor/rules/actuary-sme-protocol.mdc. Run it in your agent session
# AFTER `just preflight` exits 0.
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
    @cd annuity_model && python scripts/check_documentation_map.py
    @echo ""
    @echo "READY TO COMMIT: all four canonical gates passed."
    @echo "Reminder: trigger gate 5 from the AI agent -- '!actuaryreview'"
    @echo "or a natural-language 'actuary review' request -- when the"
    @echo "session edited any calculation / tolerance file."

# Portfolio end-to-end acceptance (extends preflight; requires portfolio flag
# for CLI / deep_smoke portfolio step). Add as a required status check when
# branch protection is updated for this workflow.
# Ring 7 acceptance (plan): preflight + portfolio parity + integration +
# deep_smoke (with portfolio) + contract checks + CLI golden + Gate 5 evidence
# (`actuary-review-full`) + reminder for manual Excel ModelCheck on portfolio.xlsx.
portfolio-acceptance:
    @just preflight
    @echo "[ring7 2/8] pytest tests/parity/portfolio"
    @cd annuity_model && ANNUITY_MODEL_PORTFOLIO_V1=1 python -m pytest tests/parity/portfolio -q
    @echo "[ring7 3/8] pytest tests/integration"
    @cd annuity_model && ANNUITY_MODEL_PORTFOLIO_V1=1 python -m pytest tests/integration -q
    @echo "[ring7 4/8] deep_smoke (portfolio workbook when ANNUITY_MODEL_PORTFOLIO_V1=1)"
    @cd annuity_model && ANNUITY_MODEL_PORTFOLIO_V1=1 python scripts/deep_smoke.py
    @echo "[ring7 5/8] render_parity_contract --check"
    @cd annuity_model && python scripts/render_parity_contract.py --check
    @echo "[ring7 6/8] CLI portfolio-run vs golden portfolio_summary.json"
    @cd annuity_model && rm -rf .smoke/portfolio_acceptance && ANNUITY_MODEL_PORTFOLIO_V1=1 python -m cli portfolio-run --inforce tests/data/inforce/example_v1/inforce.csv --out .smoke/portfolio_acceptance/
    @cd annuity_model && python -c 'import json, pathlib; root = pathlib.Path("tests/data/inforce/example_v1/expected_summary.json"); got = pathlib.Path(".smoke/portfolio_acceptance/portfolio_summary.json"); assert json.loads(root.read_text()) == json.loads(got.read_text()), "portfolio_summary.json drift"'
    @echo "[ring7 7/8] Gate 5 deterministic evidence (full scope)"
    @just actuary-review-full
    @echo "[ring7 8/8] Manual: inspect .smoke/portfolio_acceptance/portfolio.xlsx ModelCheck links and validator output (AGENTS.md)."
    @echo "RING 7 + GATE 5 (deterministic): complete. Narrative SME verdict: see .cursor/actuary-reviews/iter-1-full-ring7-gate5.md"

# Generate the Actuary SME evidence pack for an incremental review (the
# deterministic half of gate 5). The narrative subagent half is invoked
# from the AI agent via `!actuaryreview` per .cursor/rules/actuary-sme-protocol.mdc.
actuary-review:
    @cd annuity_model && python scripts/run_actuary_review.py \
        --scope incremental --iteration 1

# Same as `actuary-review` but full-project scope (used at end-of-phase
# or when the change has cross-cutting impact).
actuary-review-full:
    @cd annuity_model && python scripts/run_actuary_review.py \
        --scope full --iteration 1

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
