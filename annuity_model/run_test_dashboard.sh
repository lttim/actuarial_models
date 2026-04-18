#!/usr/bin/env bash
# macOS / Linux launcher for test_dashboard.py.
# Windows users: run_test_dashboard.bat does the same thing.
set -euo pipefail
cd "$(dirname "$0")"

if command -v python3 >/dev/null 2>&1; then
    PY=python3
else
    PY=python
fi

exec "$PY" -m streamlit run test_dashboard.py
