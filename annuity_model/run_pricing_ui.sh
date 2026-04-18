#!/usr/bin/env bash
# macOS / Linux launcher for pricing_ui.py.
# Windows: double-click run_pricing_ui.bat
# macOS:   double-click run_pricing_ui.command (opens Terminal + Chrome)
#
# Hardening rules enforced here (mirrored by tests/test_launcher_invariants.py):
#   1. PROJECT-VENV-FIRST: if ./.venv/bin/python exists, use it.
#   2. MIN-PYTHON: require Python >= MIN_PYTHON (kept in sync with pyproject.toml).
#   3. IMPORT-SMOKE: confirm `pricing_ui` itself imports before launching streamlit.
#   4. SELF-CHECK: `--self-check` runs (1)+(2)+(3) and exits 0/non-zero without
#      starting streamlit -- used by CI to catch launcher regressions early.
set -euo pipefail

# Keep MIN_PYTHON in sync with pyproject.toml [project].requires-python.
# A meta-test asserts both stay aligned.
MIN_PYTHON_MAJOR=3
MIN_PYTHON_MINOR=11

cd "$(dirname "$0")"

SELF_CHECK=0
if [[ "${1:-}" == "--self-check" ]]; then
    SELF_CHECK=1
fi

# 1. Pick interpreter.
#    Prefer the project's own venv (Python + pinned deps), then the active
#    venv, then python3 / python on PATH. We intentionally do NOT silently
#    fall back to system python if the project venv exists -- that is what
#    let a stale Python 3.9 run the app in past incidents.
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

Fix options:
  1. Install Python ${MIN_PYTHON_MAJOR}.${MIN_PYTHON_MINOR}+ from https://www.python.org/downloads/macos/
     (or `brew install python@3.12` on Homebrew).
  2. Create the project venv:
       python3 -m venv ./.venv
       ./.venv/bin/python -m pip install -r requirements.txt
  3. Activate your venv first, then re-run this script.

EOF
    exit 1
fi

# 2. Enforce minimum Python version. (slots=True, walrus, etc. require >=3.10;
#    the project's typing surface targets 3.11.)
if ! "$PY" - <<PYEOF
import sys
required = (${MIN_PYTHON_MAJOR}, ${MIN_PYTHON_MINOR})
sys.exit(0 if sys.version_info[:2] >= required else 1)
PYEOF
then
    actual_ver="$("$PY" -c 'import sys; print(".".join(str(n) for n in sys.version_info[:3]))' 2>/dev/null || echo unknown)"
    cat <<EOF >&2

[ERROR] $PY is Python ${actual_ver}; this project requires >= ${MIN_PYTHON_MAJOR}.${MIN_PYTHON_MINOR}.

Fix options:
  1. Create the project venv with a newer Python:
       python3.12 -m venv ./.venv
       ./.venv/bin/python -m pip install -r requirements.txt
     Then re-run this launcher; it picks up ./.venv automatically.
  2. Or install Python ${MIN_PYTHON_MAJOR}.${MIN_PYTHON_MINOR}+ from https://www.python.org/downloads/macos/
     (or `brew install python@3.12`) and remove ./.venv to start fresh.

EOF
    exit 1
fi

# 3. Make sure runtime deps are present. We only auto-install when running in
#    the project venv -- never into the system Python (PEP 668).
if ! "$PY" -c "import streamlit" >/dev/null 2>&1; then
    if [[ "$PY" == "./.venv/bin/python" || -n "${VIRTUAL_ENV:-}" ]]; then
        echo "[INFO] Installing pinned dependencies into the active environment..." >&2
        "$PY" -m pip install -r requirements.txt
    else
        cat <<EOF >&2

[ERROR] streamlit is not importable with $PY.
        Refusing to pip install into a non-venv interpreter (PEP 668).

Fix options:
  1. Create the project venv (recommended):
       python3 -m venv ./.venv
       ./.venv/bin/python -m pip install -r requirements.txt
  2. Or activate your own venv first, then re-run this script.

EOF
        exit 1
    fi
fi

# 4. Import-smoke the project itself. Catches regressions where deps install
#    fine but the project module tree won't load (e.g. a syntax-feature mismatch
#    like dataclass(slots=True) on Python 3.9).
if ! "$PY" -c "import pricing_ui" >/dev/null 2>&1; then
    echo "[ERROR] Failed to import pricing_ui with $PY. Re-running with full traceback:" >&2
    "$PY" -c "import pricing_ui" || true
    cat <<EOF >&2

This usually means:
  - The interpreter is too old for the project (see MIN-PYTHON above), OR
  - A required dependency is missing/incompatible, OR
  - A code change broke import-time behaviour.

Recreate the project venv to get a known-good environment:
    rm -rf ./.venv
    python3.12 -m venv ./.venv
    ./.venv/bin/python -m pip install -r requirements.txt

EOF
    exit 1
fi

if [[ "$SELF_CHECK" -eq 1 ]]; then
    echo "[OK] Launcher self-check passed: $PY ($("$PY" -c 'import sys; print(".".join(str(n) for n in sys.version_info[:3]))'))"
    exit 0
fi

# macOS: open Google Chrome to the app (Streamlit stays bound to this terminal).
# Headless avoids the first-run email prompt when stdin is not a TTY.
if [[ "$(uname -s)" == "Darwin" ]]; then
    export STREAMLIT_SERVER_HEADLESS=true
    url="http://localhost:8501"
    "$PY" -m streamlit run pricing_ui.py --server.headless true --server.port 8501 &
    st_pid=$!
    trap 'kill "$st_pid" 2>/dev/null; wait "$st_pid" 2>/dev/null; exit 130' INT TERM

    for _ in $(seq 1 120); do
        if curl -sf "${url}/_stcore/health" >/dev/null 2>&1; then
            break
        fi
        if ! kill -0 "$st_pid" 2>/dev/null; then
            echo "[ERROR] Streamlit exited before the server became ready." >&2
            wait "$st_pid" || true
            exit 1
        fi
        sleep 0.25
    done

    if [[ -d "/Applications/Google Chrome.app" ]]; then
        open -a "Google Chrome" "$url"
    else
        echo "[WARN] Google Chrome not found under /Applications; using the default browser." >&2
        open "$url"
    fi

    wait "$st_pid"
fi

exec "$PY" -m streamlit run pricing_ui.py
