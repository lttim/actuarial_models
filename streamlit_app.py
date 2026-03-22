"""
Streamlit Cloud entry point (repo root).

In Streamlit Community Cloud, set Main file path to: streamlit_app.py
(Dependencies: root requirements.txt → annuity_model/requirements.txt)
"""

from __future__ import annotations

import sys
from pathlib import Path

_APP_DIR = Path(__file__).resolve().parent / "annuity_model"
if str(_APP_DIR) not in sys.path:
    sys.path.insert(0, str(_APP_DIR))

import spia_ui

spia_ui.main()
