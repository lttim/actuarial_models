#!/usr/bin/env bash
# Bootstrap the annuity_model Python environment on macOS / Linux and run the
# regression suite. Idempotent — safe to re-run.
#
# Usage:
#   cd annuity_model
#   ./bootstrap_macos.sh                # default: create .venv, install, test
#   ./bootstrap_macos.sh --no-tests     # skip the pytest gate
#
set -euo pipefail

cd "$(dirname "$0")"

run_tests=1
for arg in "$@"; do
    case "$arg" in
        --no-tests) run_tests=0 ;;
        -h|--help)
            sed -n '2,12p' "$0"
            exit 0
            ;;
        *) echo "[ERROR] Unknown argument: $arg" >&2; exit 2 ;;
    esac
done

# 1. Find a usable Python 3.11+.
pick_python() {
    for cand in python3.13 python3.12 python3.11 python3; do
        if command -v "$cand" >/dev/null 2>&1; then
            ver=$("$cand" -c 'import sys; print("%d.%d" % sys.version_info[:2])')
            major=${ver%%.*}
            minor=${ver##*.}
            if [[ "$major" -ge 3 && "$minor" -ge 11 ]]; then
                echo "$cand"
                return 0
            fi
        fi
    done
    return 1
}

if ! PY=$(pick_python); then
    cat <<'EOF' >&2

[ERROR] No Python 3.11+ found on PATH.

On macOS:
  brew install python@3.12          # then re-open this terminal
On Linux:
  sudo apt install python3.12-venv  # or your distro equivalent

EOF
    exit 1
fi

py_ver=$("$PY" -c 'import sys; print(".".join(map(str, sys.version_info[:3])))')
echo "[bootstrap] Using $PY ($py_ver)"

# 2. Create / refresh the virtualenv.
if [[ ! -d .venv ]]; then
    echo "[bootstrap] Creating .venv ..."
    "$PY" -m venv .venv
fi

# shellcheck disable=SC1091
source .venv/bin/activate

# 3. Install pinned deps.
echo "[bootstrap] Upgrading pip ..."
python -m pip install --upgrade pip >/dev/null

echo "[bootstrap] Installing requirements.txt ..."
python -m pip install -r requirements.txt

if [[ -f requirements-dev.txt ]]; then
    echo "[bootstrap] Installing requirements-dev.txt ..."
    python -m pip install -r requirements-dev.txt
fi

# 4. Smoke import — fail fast if a dep is wrong before launching pytest.
python - <<'PY'
import importlib.util
import sys

mods = ["numpy", "pandas", "openpyxl", "streamlit", "matplotlib", "pyarrow"]
missing = [m for m in mods if importlib.util.find_spec(m) is None]
if missing:
    sys.stderr.write(f"[bootstrap] Missing modules: {missing}\n")
    sys.exit(1)
print(f"[bootstrap] All core deps importable on Python {sys.version.split()[0]}")
PY

# 5. Optional regression gate.
if [[ "$run_tests" -eq 1 ]]; then
    echo "[bootstrap] Running pytest tests/ tests/parity/ ..."
    python -m pytest tests/ tests/parity/ -q
    echo
    echo "[bootstrap] OK — all tests passed."
else
    echo "[bootstrap] Skipped tests (--no-tests)."
fi

cat <<'EOF'

[bootstrap] Done. Suggested next steps:

    source .venv/bin/activate                 # if not already active
    ./run_pricing_ui.sh                       # Streamlit UI on :8501
    ./run_tests_report.sh                     # regenerate HTML pytest report
    pytest tests/test_excel_export_validation.py -v   # validate Excel exports

EOF
