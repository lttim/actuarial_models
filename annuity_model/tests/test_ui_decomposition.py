"""Focused guards for the incremental Streamlit UI decomposition."""

from __future__ import annotations

import ast
import datetime as dt
import math
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC_ROOT = ROOT / "src"


def test_navigation_options_inject_portfolio_after_pricing_run() -> None:
    code = (
        "import sys\n"
        f"sys.path.insert(0, {str(SRC_ROOT)!r})\n"
        "from annuity_model.ui.navigation import SECTION_LABELS, SECTION_ORDER, section_options\n"
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
        f"sys.path.insert(0, {str(SRC_ROOT)!r})\n"
        "from annuity_model.ui.navigation import overview_section_labels\n"
        "from annuity_model.ui.pages.overview import dynamic_overview_features\n"
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
    pricing_ui_path = SRC_ROOT / "annuity_model" / "pricing_ui.py"
    tree = ast.parse(pricing_ui_path.read_text(encoding="utf-8"))

    imported_modules = {
        node.module
        for node in ast.walk(tree)
        if isinstance(node, ast.ImportFrom) and node.module is not None
    }
    function_names = {
        node.name
        for node in ast.walk(tree)
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
    }

    assert "annuity_model.ui.app_shell" in imported_modules
    assert "annuity_model.ui.diagnostics" in imported_modules
    assert "annuity_model.ui.pages.overview" in imported_modules
    assert "annuity_model.ui.pages.router" in imported_modules
    assert "annuity_model.ui.widgets.product_badges" in imported_modules
    assert "_render_overview" not in function_names


def test_product_badges_are_product_definition_backed() -> None:
    code = (
        "import sys\n"
        f"sys.path.insert(0, {str(SRC_ROOT)!r})\n"
        "from annuity_model.product_registry import product_options_for_ui\n"
        "from annuity_model.ui.widgets.product_badges import badges_for_status, product_statuses\n"
        "statuses = product_statuses()\n"
        "assert len(statuses) == len(product_options_for_ui())\n"
        "assert all(s.maturity_label == 'Mechanics-production' for s in statuses)\n"
        "assert all(s.assumption_profile == 'demo-safe-with-waiver' for s in statuses)\n"
        "assert any('Monte Carlo' in badge for s in statuses for badge in badges_for_status(s))\n"
    )
    proc = subprocess.run(
        [sys.executable, "-c", code],
        capture_output=True,
        text=True,
        timeout=30,
        cwd=str(ROOT),
    )
    assert proc.returncode == 0, proc.stderr + proc.stdout


def test_diagnostics_payload_builder_covers_pricing_and_empty_optional_sections() -> None:
    if str(SRC_ROOT) not in sys.path:
        sys.path.insert(0, str(SRC_ROOT))
    from annuity_model.ui.diagnostics import (  # noqa: PLC0415
        DiagnosticsBuilders,
        MissingDiagnosticsInput,
        build_diagnostics_payload,
    )

    builders = DiagnosticsBuilders(
        active_provenance_rows=lambda: [{"source": "fixture"}],
        pricing_result_to_dict=lambda res, contract, include_full: {
            "res": res,
            "contract": contract,
            "include_full": include_full,
        },
        yield_curve_to_dict=lambda value: {"yield_curve": value},
        mortality_to_dict=lambda value: {"mortality": value},
        alm_result_to_dict=lambda *args, **kwargs: {"alm": True, "kwargs": kwargs},
        alm_assumptions_to_dict=lambda value: {"assumptions": value},
        whatif_result_to_dict=lambda **kwargs: {"what_if": True, "kwargs": kwargs},
        is_yield_curve=lambda value: value == "curve",
        is_expense_assumptions=lambda value: value == "expenses",
        is_alm_result=lambda value: value == "alm_result",
        is_alm_assumptions=lambda value: value == "alm_assumptions",
    )

    try:
        build_diagnostics_payload({}, builders=builders)
    except MissingDiagnosticsInput:
        pass
    else:
        raise AssertionError("missing pricing inputs should block diagnostics payload")

    payload = build_diagnostics_payload(
        {
            "pricing_res": "result",
            "pricing_contract": "contract",
            "pricing_run_id": "run-1",
            "pricing_meta": {"product_type": "spia"},
            "pricing_excel_context": {
                "yield_curve": "curve",
                "mortality": "mortality",
                "expenses": "expenses",
                "yield_mode": "flat",
            },
        },
        builders=builders,
        exported_at_utc=dt.datetime(2026, 5, 10, 12, 0, 0),
    )

    assert payload["exported_at_utc"] == "2026-05-10T12:00:00Z"
    assert payload["pricing"]["include_full"] is True
    assert payload["pricing_inputs"]["yield_curve"] == {"yield_curve": "curve"}
    assert math.isnan(payload["pricing_inputs"]["expenses"]["premium_expense_rate"])
    assert payload["assumption_provenance"] == [{"source": "fixture"}]
    assert payload["alm"] is None
    assert payload["what_if"] is None
