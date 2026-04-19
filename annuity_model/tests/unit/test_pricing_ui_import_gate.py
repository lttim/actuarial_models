"""Gate imports for ``pricing_ui`` — catches broken cross-module exports early.

``pricing_ui`` is not imported by the default unit-test matrix as a side effect of
other tests. A bad ``from portfolio_config import …`` (or similar) therefore only
surfaced when someone ran ``streamlit run pricing_ui.py``. These tests fail in CI
on the same class of mistake.
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[2]  # …/annuity_model/


def test_pricing_ui_import_succeeds_in_clean_subprocess() -> None:
    """Fresh interpreter: full ``pricing_ui`` import must not raise (e.g. ImportError)."""
    code = (
        "import importlib.util\n"
        "import sys\n"
        f"sys.path.insert(0, {str(ROOT)!r})\n"
        "importlib.import_module('pricing_ui')\n"
    )
    proc = subprocess.run(
        [sys.executable, "-c", code],
        capture_output=True,
        text=True,
        timeout=120,
        cwd=str(ROOT),
    )
    assert proc.returncode == 0, proc.stderr + proc.stdout


def test_portfolio_config___all___exports_are_defined() -> None:
    """Every name in ``portfolio_config.__all__`` must exist (prevents stale __all__)."""
    import portfolio_config as m

    for name in m.__all__:
        assert hasattr(m, name), f"portfolio_config.__all__ lists missing name: {name!r}"


def test_portfolio_v1_enabled_avoids_streamlit_until_sidebar_visible() -> None:
    """CLI entrypoints import ``portfolio_config`` without pulling Streamlit."""
    code = (
        "import sys\n"
        f"sys.path.insert(0, {str(ROOT)!r})\n"
        "import portfolio_config as p\n"
        "assert 'streamlit' not in sys.modules\n"
        "p.portfolio_v1_enabled()\n"
        "assert 'streamlit' not in sys.modules\n"
    )
    proc = subprocess.run(
        [sys.executable, "-c", code],
        capture_output=True,
        text=True,
        timeout=30,
        cwd=str(ROOT),
    )
    assert proc.returncode == 0, proc.stderr + proc.stdout


@pytest.mark.parametrize(
    "symbol",
    (
        "portfolio_disable_file_path",
        "portfolio_disabled_explanation_markdown",
        "portfolio_sidebar_visible",
        "portfolio_v1_enabled",
    ),
)
def test_pricing_ui_expected_portfolio_config_imports(symbol: str) -> None:
    """Names ``pricing_ui`` imports from ``portfolio_config`` must stay importable."""
    import importlib

    m = importlib.import_module("portfolio_config")
    assert hasattr(m, symbol), f"portfolio_config.{symbol} missing"
