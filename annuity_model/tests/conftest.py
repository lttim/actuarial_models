"""Pytest hooks shared by the whole ``tests/`` tree."""

from __future__ import annotations

import sys

import pytest


def pytest_configure(config: pytest.Config) -> None:
    """Fail fast with a clear message when CI or a developer uses an unsupported interpreter."""
    from annuity_model.pytest_python import MIN_PYTHON_MAJOR, MIN_PYTHON_MINOR

    if sys.version_info[:2] < (MIN_PYTHON_MAJOR, MIN_PYTHON_MINOR):
        pytest.exit(
            f"This suite requires Python >= {MIN_PYTHON_MAJOR}.{MIN_PYTHON_MINOR} "
            "(see `pyproject.toml` and `pytest_python.py`). "
            "Create `./.venv` with Python 3.11+ or run `./run_pricing_ui.sh` / `./run_test_dashboard.sh`.",
            returncode=4,
        )
