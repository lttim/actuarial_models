#!/usr/bin/env bash
# macOS / Linux pytest-html launcher.
# Windows users: run_tests_report.bat does the same thing.
set -euo pipefail
cd "$(dirname "$0")"

mkdir -p reports

if command -v python3 >/dev/null 2>&1; then
    PY=python3
else
    PY=python
fi

"$PY" -m pytest --html=reports/pytest_report.html --self-contained-html
status=$?

# Open the report if a GUI opener is available (macOS: open; Linux: xdg-open).
if [[ -f reports/pytest_report.html ]]; then
    if command -v open >/dev/null 2>&1; then
        open reports/pytest_report.html || true
    elif command -v xdg-open >/dev/null 2>&1; then
        xdg-open reports/pytest_report.html >/dev/null 2>&1 || true
    fi
fi

exit "$status"
