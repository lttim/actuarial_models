"""
Pick a Python interpreter for subprocess test runs (pytest, CI helpers).

Keeps `test_dashboard.py`, launchers, and meta-tests aligned on the same rules:
prefer the project ``.venv``, enforce ``pyproject.toml`` ``requires-python``, and
never silently fall back to an ancient system Python when a stale ``.venv``
exists (the original Streamlit + pytest incident class).
"""

from __future__ import annotations

import os
import re
import subprocess
import sys
import tomllib
from pathlib import Path

_ANCHOR = Path(__file__).resolve().parent


def min_python_from_pyproject(anchor: Path = _ANCHOR) -> tuple[int, int]:
    """Parse ``[project].requires-python`` (expects ``>=M.m``)."""
    path = anchor / "pyproject.toml"
    data = tomllib.loads(path.read_text(encoding="utf-8"))
    spec = data["project"]["requires-python"]
    m = re.match(r"^>=\s*(\d+)\.(\d+)\s*$", spec)
    if not m:
        raise ValueError(f"Unsupported requires-python form: {spec!r}")
    return int(m.group(1)), int(m.group(2))


MIN_PYTHON_MAJOR, MIN_PYTHON_MINOR = min_python_from_pyproject()


def project_venv_python(anchor: Path) -> Path | None:
    """Return the project venv interpreter path if that file exists."""
    if os.name == "nt":
        p = anchor / ".venv" / "Scripts" / "python.exe"
    else:
        p = anchor / ".venv" / "bin" / "python"
    return p if p.is_file() else None


def interpreter_meets_minimum(py: str | Path, major: int, minor: int) -> bool:
    proc = subprocess.run(
        [str(py), "-c", f"import sys; raise SystemExit(0 if sys.version_info[:2] >= ({major}, {minor}) else 1)"],
        capture_output=True,
    )
    return proc.returncode == 0


def select_pytest_interpreter(anchor: Path | None = None) -> tuple[str | None, str | None]:
    """
    Choose ``python`` for ``subprocess``-driven pytest from the annuity_model tree.

    ``anchor`` is the directory that may contain ``.venv`` (usually ``annuity_model/``).
    Minimum Python is always read from ``pyproject.toml`` beside this module.

    Returns ``(executable, None)`` on success, or ``(None, human_message)`` on failure.
    """
    root = anchor.resolve() if anchor is not None else _ANCHOR
    req = min_python_from_pyproject(_ANCHOR)
    venv_py = project_venv_python(root)

    if venv_py is not None:
        if interpreter_meets_minimum(venv_py, *req):
            return str(venv_py), None
        return None, (
            f"The project virtualenv at `{venv_py}` uses a Python older than "
            f"{req[0]}.{req[1]} (see `pyproject.toml`). Remove it and recreate, e.g.\n"
            f"  rm -rf .venv && python3.12 -m venv .venv && .venv/bin/python -m pip install -r requirements.txt -r requirements-dev.txt"
        )

    if interpreter_meets_minimum(sys.executable, *req):
        return sys.executable, None

    return None, (
        f"No usable interpreter: need Python >= {req[0]}.{req[1]} (see `pyproject.toml`). "
        "Create `./.venv` with Python 3.11+ and `pip install -r requirements.txt -r requirements-dev.txt`, "
        "or launch Streamlit via `./run_pricing_ui.sh` / `./run_test_dashboard.sh` so the venv is picked up."
    )
