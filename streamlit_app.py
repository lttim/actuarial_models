"""
Streamlit Cloud entry point (repo root).

In Streamlit Community Cloud, set Main file path to: streamlit_app.py
"""

from __future__ import annotations

import os

# Headless Linux (Streamlit Cloud): avoid matplotlib trying a GUI backend if it loads.
os.environ.setdefault("MPLBACKEND", "Agg")

import streamlit as st


def _launch() -> None:
    import sys
    from pathlib import Path

    app_dir = Path(__file__).resolve().parent / "annuity_model"
    if str(app_dir) not in sys.path:
        sys.path.insert(0, str(app_dir))

    import spia_ui

    spia_ui.main()


try:
    _launch()
except Exception as exc:
    try:
        st.set_page_config(page_title="SPIA workspace — startup error", layout="wide")
    except Exception:
        pass
    st.error("The app failed while starting. Details below (copy for debugging).")
    st.exception(exc)
