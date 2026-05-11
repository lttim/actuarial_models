"""Unit tests for ``pytest_python`` (interpreter selection for subprocess pytest)."""

from __future__ import annotations

import os
import sys
from pathlib import Path

import pytest

from annuity_model import pytest_python as pp
from annuity_model.pytest_python import select_pytest_interpreter

PACKAGE_ROOT = Path(__file__).resolve().parent.parent


def test_min_python_constants_match_pyproject() -> None:
    assert pp.min_python_from_pyproject(PACKAGE_ROOT) == (pp.MIN_PYTHON_MAJOR, pp.MIN_PYTHON_MINOR)


def test_project_venv_python_none_for_missing_venv(tmp_path: Path) -> None:
    assert pp.project_venv_python(tmp_path) is None


def test_select_pytest_interpreter_uses_sys_executable_without_venv(tmp_path: Path) -> None:
    """With no ``.venv`` under ``anchor``, a new-enough ``sys.executable`` is used."""
    exe, err = select_pytest_interpreter(tmp_path)
    if sys.version_info[:2] >= (pp.MIN_PYTHON_MAJOR, pp.MIN_PYTHON_MINOR):
        assert err is None
        assert exe == sys.executable
    else:
        assert exe is None
        assert err is not None


def test_select_pytest_interpreter_prefers_project_venv(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """When a project venv interpreter exists and meets the floor, it is chosen over ``sys.executable``."""

    if sys.version_info[:2] < (pp.MIN_PYTHON_MAJOR, pp.MIN_PYTHON_MINOR):
        pytest.skip("runner Python below project minimum")

    resolved = tmp_path.resolve()

    def fake_venv(anchor: Path) -> Path | None:
        return Path(sys.executable) if anchor.resolve() == resolved else None

    monkeypatch.setattr(pp, "project_venv_python", fake_venv)
    exe, err = select_pytest_interpreter(tmp_path)
    assert err is None
    assert exe == sys.executable


def test_select_pytest_interpreter_rejects_interpreter_without_pytest(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """A Python that meets the version floor is still unusable for pytest commands without pytest."""
    monkeypatch.setattr(pp, "project_venv_python", lambda anchor: None)
    monkeypatch.setattr(pp, "interpreter_meets_minimum", lambda py, maj, min_: True)
    monkeypatch.setattr(pp, "interpreter_has_pytest", lambda py: False)

    exe, err = select_pytest_interpreter(tmp_path)

    assert exe is None
    assert err is not None
    assert "pytest" in err
    assert "requirements-dev.txt" in err


def test_select_pytest_interpreter_rejects_stale_venv(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Stale ``.venv`` with a too-old interpreter must error (no silent fallback)."""
    monkeypatch.setattr(pp, "interpreter_meets_minimum", lambda py, maj, min_: False)
    if os.name == "nt":
        venv_bin = tmp_path / ".venv" / "Scripts"
        py_name = "python.exe"
    else:
        venv_bin = tmp_path / ".venv" / "bin"
        py_name = "python"
    venv_bin.mkdir(parents=True)
    (venv_bin / py_name).write_text("")
    (venv_bin / py_name).chmod(0o755)

    exe, err = select_pytest_interpreter(tmp_path)
    assert exe is None
    assert err is not None
    assert ".venv" in err or "virtualenv" in err.lower()
