#!/bin/bash
# Double-click this file in Finder: opens Terminal, starts Streamlit, opens Chrome.
cd "$(dirname "$0")" || exit 1
exec bash ./run_pricing_ui.sh
