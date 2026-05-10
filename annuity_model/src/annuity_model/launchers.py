"""Console entry points for local Streamlit applications."""

from __future__ import annotations

import subprocess  # nosec B404
import sys
from pathlib import Path


def _run_streamlit(script_name: str) -> int:
    script = Path(__file__).resolve().with_name(script_name)
    # Reviewed: fixed Streamlit command vector; argv is passed through as user CLI flags.
    return subprocess.call(  # nosec B603
        [sys.executable, "-m", "streamlit", "run", str(script), *sys.argv[1:]]
    )


def pricing_ui() -> int:
    """Launch the pricing Streamlit UI from an installed package."""
    return _run_streamlit("pricing_ui.py")


def test_dashboard() -> int:
    """Launch the embedded test dashboard from an installed package."""
    return _run_streamlit("test_dashboard.py")
