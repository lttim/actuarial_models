#!/usr/bin/env bash
# macOS / Linux launcher for test_dashboard.py (embedded-style dashboard only).
# Windows: run_test_dashboard.bat
#
# Hardening mirrors run_pricing_ui.sh (see tests/test_launcher_invariants.py):
#   1. PROJECT-VENV-FIRST
#   2. MIN-PYTHON from pyproject.toml
#   3. IMPORT-SMOKE: ``import annuity_model.test_dashboard`` before Streamlit
#   4. ``--self-check`` runs (1)+(2)+(3) without starting the server
set -euo pipefail

MIN_PYTHON_MAJOR=3
MIN_PYTHON_MINOR=11

cd "$(dirname "$0")"

export PYTHONPATH="$PWD/src:${PYTHONPATH:-}"
APP_SCRIPT="src/annuity_model/test_dashboard.py"

SELF_CHECK=0
if [[ "${1:-}" == "--self-check" ]]; then
    SELF_CHECK=1
fi

PY=""
if [[ -x "./.venv/bin/python" ]]; then
    PY="./.venv/bin/python"
elif [[ -n "${VIRTUAL_ENV:-}" && -x "${VIRTUAL_ENV}/bin/python" ]]; then
    PY="${VIRTUAL_ENV}/bin/python"
elif command -v python3 >/dev/null 2>&1; then
    PY="$(command -v python3)"
elif command -v python >/dev/null 2>&1; then
    PY="$(command -v python)"
else
    cat <<EOF >&2
[ERROR] No usable Python interpreter on PATH (need >= ${MIN_PYTHON_MAJOR}.${MIN_PYTHON_MINOR}).
EOF
    exit 1
fi

if ! "$PY" - <<PYEOF
import sys
required = (${MIN_PYTHON_MAJOR}, ${MIN_PYTHON_MINOR})
sys.exit(0 if sys.version_info[:2] >= required else 1)
PYEOF
then
    cat <<EOF >&2
[ERROR] $PY is too old; this project requires >= ${MIN_PYTHON_MAJOR}.${MIN_PYTHON_MINOR}.
EOF
    exit 1
fi

if ! "$PY" -c "import streamlit" >/dev/null 2>&1; then
    if [[ "$PY" == "./.venv/bin/python" || -n "${VIRTUAL_ENV:-}" ]]; then
        echo "[INFO] Installing pinned dependencies..." >&2
        "$PY" -m pip install -r requirements.txt
    else
        cat <<EOF >&2
[ERROR] streamlit not importable; create ./.venv and pip install -r requirements.txt
EOF
        exit 1
    fi
fi

if ! "$PY" -c "import annuity_model.test_dashboard" >/dev/null 2>&1; then
    echo "[ERROR] Failed to import annuity_model.test_dashboard with $PY:" >&2
    "$PY" -c "import annuity_model.test_dashboard" || true
    exit 1
fi

if [[ "$SELF_CHECK" -eq 1 ]]; then
    echo "[OK] test_dashboard launcher self-check passed: $PY ($("$PY" -c 'import sys; print(".".join(str(n) for n in sys.version_info[:3]))'))"
    exit 0
fi

exec "$PY" -m streamlit run "$APP_SCRIPT"
