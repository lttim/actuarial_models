#!/usr/bin/env bash
# macOS / Linux launcher for pricing_ui.py.
# Windows users: run_pricing_ui.bat does the same thing.
set -euo pipefail

cd "$(dirname "$0")"

# Prefer an active virtualenv; otherwise fall back to python3, then python.
if [[ -n "${VIRTUAL_ENV:-}" ]] && command -v python >/dev/null 2>&1; then
    PY=python
elif command -v python3 >/dev/null 2>&1; then
    PY=python3
elif command -v python >/dev/null 2>&1; then
    PY=python
else
    cat <<'EOF' >&2

[ERROR] No usable Python interpreter on PATH.

Fix options:
  1. Install Python 3.11+ from https://www.python.org/downloads/macos/
     (or `brew install python@3.12` on Homebrew).
  2. Activate your venv first, then re-run this script.

EOF
    exit 1
fi

# Make sure Streamlit is available; surface a helpful hint if not.
if ! "$PY" -c "import streamlit" >/dev/null 2>&1; then
    echo "[INFO] Installing pinned dependencies into the active environment..." >&2
    "$PY" -m pip install -r requirements.txt
fi

exec "$PY" -m streamlit run pricing_ui.py
