"""Focused guards for the incremental Streamlit UI decomposition."""

from __future__ import annotations

import ast
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def test_navigation_options_inject_portfolio_after_pricing_run() -> None:
    code = (
        "import sys\n"
        f"sys.path.insert(0, {str(ROOT)!r})\n"
        "from ui.navigation import SECTION_LABELS, SECTION_ORDER, section_options\n"
        "assert section_options(include_portfolio=False) == list(SECTION_ORDER)\n"
        "with_portfolio = section_options(include_portfolio=True)\n"
        "assert with_portfolio[:3] == ['overview', 'run', 'portfolio']\n"
        "assert with_portfolio[3:] == list(SECTION_ORDER[2:])\n"
        "assert SECTION_LABELS['portfolio'] == 'Portfolio (multi-policy)'\n"
    )
    proc = subprocess.run(
        [sys.executable, "-c", code],
        capture_output=True,
        text=True,
        timeout=30,
        cwd=str(ROOT),
    )
    assert proc.returncode == 0, proc.stderr + proc.stdout


def test_overview_features_are_registry_backed() -> None:
    code = (
        "import sys\n"
        f"sys.path.insert(0, {str(ROOT)!r})\n"
        "from ui.navigation import overview_section_labels\n"
        "from ui.pages.overview import dynamic_overview_features\n"
        "features = dynamic_overview_features()\n"
        "assert len(features) >= 8\n"
        "assert any('Supported product run types' in feature for feature in features)\n"
        "assert any('Monte Carlo pricing enabled' in feature for feature in features)\n"
        "assert 'Pricing Run' in overview_section_labels()\n"
    )
    proc = subprocess.run(
        [sys.executable, "-c", code],
        capture_output=True,
        text=True,
        timeout=30,
        cwd=str(ROOT),
    )
    assert proc.returncode == 0, proc.stderr + proc.stdout


def test_pricing_ui_delegates_overview_and_shell_to_ui_modules() -> None:
    pricing_ui_path = ROOT / "pricing_ui.py"
    tree = ast.parse(pricing_ui_path.read_text(encoding="utf-8"))

    imported_modules = {
        node.module
        for node in ast.walk(tree)
        if isinstance(node, ast.ImportFrom) and node.module is not None
    }
    function_names = {
        node.name for node in ast.walk(tree) if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
    }

    assert "annuity_model.ui.app_shell" in imported_modules
    assert "annuity_model.ui.pages.overview" in imported_modules
    assert "_render_overview" not in function_names
