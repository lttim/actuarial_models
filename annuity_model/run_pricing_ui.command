#!/bin/bash
# Double-click this file in Finder: opens Terminal, starts Streamlit, opens Chrome.
# All real logic lives in run_pricing_ui.sh; this wrapper exists so Finder can
# launch it. We deliberately keep Terminal open on failure so the user can read
# the error before the window auto-closes ("Settings > Profiles > Shell > When
# the shell exits = Close if the shell exited cleanly" is the macOS default).
cd "$(dirname "$0")" || exit 1

bash ./run_pricing_ui.sh "$@"
status=$?

if [[ "$status" -ne 0 ]]; then
    echo
    echo "[run_pricing_ui] exited with status $status."
    echo "Press Return to close this window..."
    read -r _ || true
fi

exit "$status"
