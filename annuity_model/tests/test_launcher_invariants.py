"""Launcher invariants -- end-user double-click path is part of the platform.

A regression here ships a broken product to whoever double-clicks
``run_pricing_ui.command`` from Finder, even when every parity / unit / smoke
test is green. These checks lock the launcher contract in place:

  1. The minimum Python version is declared in **one** place
     (``pyproject.toml [project].requires-python``) and the shell + batch
     launchers reference the same major.minor pair.
  2. Each launcher prefers the project ``.venv`` before falling back to system
     Python -- otherwise a stale interpreter (e.g. macOS-bundled 3.9) silently
     picks up stray site-packages and runs the app to a hard crash.
  3. Each launcher import-smokes ``pricing_ui`` (not just ``streamlit``) so a
     code-level regression that breaks module load is caught before Streamlit
     is launched and Terminal closes the window.
  4. The bash launcher's ``--self-check`` mode actually works on this machine
     with the project ``.venv`` -- this is what CI runs as a smoke test.

If a check fails, the fix lives in the offending launcher / pyproject -- not
this test.
"""

from __future__ import annotations

import os
import re
import subprocess
import sys
import tomllib
from pathlib import Path

import pytest

PACKAGE_ROOT = Path(__file__).resolve().parent.parent
PYPROJECT = PACKAGE_ROOT / "pyproject.toml"
SH = PACKAGE_ROOT / "run_pricing_ui.sh"
CMD = PACKAGE_ROOT / "run_pricing_ui.command"
BAT = PACKAGE_ROOT / "run_pricing_ui.bat"
TEST_DASH_SH = PACKAGE_ROOT / "run_test_dashboard.sh"
TEST_DASH_BAT = PACKAGE_ROOT / "run_test_dashboard.bat"

pytestmark = pytest.mark.invariant


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _requires_python_from_pyproject() -> tuple[int, int]:
    data = tomllib.loads(PYPROJECT.read_text(encoding="utf-8"))
    spec = data["project"]["requires-python"]
    m = re.match(r"^>=\s*(\d+)\.(\d+)\s*$", spec)
    assert m, (
        f"pyproject.toml [project].requires-python must look like '>=3.11', got {spec!r}. "
        "The launcher meta-tests parse this exact form."
    )
    return int(m.group(1)), int(m.group(2))


# ---------------------------------------------------------------------------
# 1. requires-python is declared and parseable.
# ---------------------------------------------------------------------------


def test_pyproject_declares_requires_python() -> None:
    major, minor = _requires_python_from_pyproject()
    assert (major, minor) >= (3, 10), (
        "Project uses dataclass(slots=True) and PEP 604 unions, which need "
        f"Python >= 3.10. requires-python pins {major}.{minor}."
    )


# ---------------------------------------------------------------------------
# 2. Shell + batch launchers pin the SAME min Python as pyproject.
# ---------------------------------------------------------------------------


def test_shell_launcher_min_python_matches_pyproject() -> None:
    major, minor = _requires_python_from_pyproject()
    text = SH.read_text(encoding="utf-8")
    assert f"MIN_PYTHON_MAJOR={major}" in text, (
        "run_pricing_ui.sh must declare MIN_PYTHON_MAJOR matching pyproject."
    )
    assert f"MIN_PYTHON_MINOR={minor}" in text, (
        "run_pricing_ui.sh must declare MIN_PYTHON_MINOR matching pyproject."
    )


def test_batch_launcher_min_python_matches_pyproject() -> None:
    major, minor = _requires_python_from_pyproject()
    text = BAT.read_text(encoding="utf-8")
    assert f"MIN_PY_MAJOR={major}" in text, (
        "run_pricing_ui.bat must declare MIN_PY_MAJOR matching pyproject."
    )
    assert f"MIN_PY_MINOR={minor}" in text, (
        "run_pricing_ui.bat must declare MIN_PY_MINOR matching pyproject."
    )


def test_test_dashboard_shell_min_python_matches_pyproject() -> None:
    major, minor = _requires_python_from_pyproject()
    text = TEST_DASH_SH.read_text(encoding="utf-8")
    assert f"MIN_PYTHON_MAJOR={major}" in text
    assert f"MIN_PYTHON_MINOR={minor}" in text


def test_test_dashboard_batch_min_python_matches_pyproject() -> None:
    major, minor = _requires_python_from_pyproject()
    text = TEST_DASH_BAT.read_text(encoding="utf-8")
    assert f"MIN_PY_MAJOR={major}" in text
    assert f"MIN_PY_MINOR={minor}" in text


TEST_DASH_SHELL_REQUIRED_CLAUSES = {
    "prefers project venv": r'\./\.venv/bin/python',
    "version guard": r"sys\.version_info\[:2\]\s*>=\s*required",
    "imports test_dashboard": r"import test_dashboard",
    "self-check mode": r"--self-check",
}


@pytest.mark.parametrize("label,pattern", list(TEST_DASH_SHELL_REQUIRED_CLAUSES.items()))
def test_test_dashboard_shell_has_required_clause(label: str, pattern: str) -> None:
    text = TEST_DASH_SH.read_text(encoding="utf-8")
    assert re.search(pattern, text), f"run_test_dashboard.sh missing '{label}' (/{pattern}/)"


TEST_DASH_BATCH_REQUIRED_CLAUSES = {
    "prefers project venv": r"\.venv\\Scripts\\python\.exe",
    "version guard": r"sys\.version_info\[:2\]\s*>=\s*\(%MIN_PY_MAJOR%,\s*%MIN_PY_MINOR%\)",
    "imports test_dashboard": r"import test_dashboard",
    "self-check mode": r"--self-check",
}


@pytest.mark.parametrize("label,pattern", list(TEST_DASH_BATCH_REQUIRED_CLAUSES.items()))
def test_test_dashboard_batch_has_required_clause(label: str, pattern: str) -> None:
    text = TEST_DASH_BAT.read_text(encoding="utf-8")
    assert re.search(pattern, text), f"run_test_dashboard.bat missing '{label}' (/{pattern}/)"


# ---------------------------------------------------------------------------
# 3. Required hardening clauses are present in each launcher.
#    These are textual checks (not behavioural) so they run on any platform.
# ---------------------------------------------------------------------------


SHELL_REQUIRED_CLAUSES = {
    "prefers project venv": r'-x\s+"\./\.venv/bin/python"',
    "version guard": r"sys\.version_info\[:2\]\s*>=\s*required",
    "imports pricing_ui (not just streamlit)": r'"\$PY"\s+-c\s+"import pricing_ui"',
    "self-check mode": r'"\$\{1:-\}"\s*==\s*"--self-check"',
    "refuses pip into system python": r"PEP 668",
}


@pytest.mark.parametrize("label,pattern", list(SHELL_REQUIRED_CLAUSES.items()))
def test_shell_launcher_has_required_clause(label: str, pattern: str) -> None:
    text = SH.read_text(encoding="utf-8")
    assert re.search(pattern, text), (
        f"run_pricing_ui.sh is missing the '{label}' guard "
        f"(no match for /{pattern}/). See AGENTS.md for the launcher contract."
    )


BATCH_REQUIRED_CLAUSES = {
    "prefers project venv": r"\.venv\\Scripts\\python\.exe",
    "version guard": r"sys\.version_info\[:2\]\s*>=\s*\(%MIN_PY_MAJOR%,\s*%MIN_PY_MINOR%\)",
    "imports pricing_ui (not just streamlit)": r"import pricing_ui",
    "self-check mode": r"--self-check",
}


@pytest.mark.parametrize("label,pattern", list(BATCH_REQUIRED_CLAUSES.items()))
def test_batch_launcher_has_required_clause(label: str, pattern: str) -> None:
    text = BAT.read_text(encoding="utf-8")
    assert re.search(pattern, text), (
        f"run_pricing_ui.bat is missing the '{label}' guard "
        f"(no match for /{pattern}/). See AGENTS.md for the launcher contract."
    )


def test_command_launcher_holds_terminal_on_error() -> None:
    """The .command wrapper must keep Terminal open on non-zero exit so the
    user can read the error before macOS auto-closes the window."""
    text = CMD.read_text(encoding="utf-8")
    assert "read -r" in text, (
        "run_pricing_ui.command must `read` after a non-zero status so the "
        "Terminal window stays open long enough to read the error."
    )
    assert "status=$?" in text, "run_pricing_ui.command must capture the launcher exit status."


# ---------------------------------------------------------------------------
# 4. Executable bits (POSIX only).
# ---------------------------------------------------------------------------


@pytest.mark.skipif(os.name == "nt", reason="POSIX exec bits not meaningful on Windows")
def test_posix_launchers_are_executable() -> None:
    for path in (SH, CMD, TEST_DASH_SH):
        assert os.access(path, os.X_OK), (
            f"{path.name} must be executable (chmod +x). Finder refuses to "
            "double-click a non-executable .command file."
        )


# ---------------------------------------------------------------------------
# 5. End-to-end: --self-check actually runs cleanly with a stripped PATH.
#    This is the same invocation CI uses; it catches the original incident
#    (system Python 3.9 found before the project venv) directly.
# ---------------------------------------------------------------------------


@pytest.mark.skipif(os.name == "nt", reason="bash launcher is POSIX-only")
def test_shell_launcher_self_check_with_clean_path() -> None:
    if not (PACKAGE_ROOT / ".venv" / "bin" / "python").exists():
        pytest.skip("project .venv not built; create it via `python3.12 -m venv .venv`")
    env = {
        "HOME": os.environ.get("HOME", "/tmp"),
        "PATH": "/usr/bin:/bin:/usr/sbin:/sbin",
        "SHELL": "/bin/bash",
    }
    result = subprocess.run(
        ["bash", str(SH), "--self-check"],
        cwd=str(PACKAGE_ROOT),
        env=env,
        capture_output=True,
        text=True,
        timeout=120,
    )
    assert result.returncode == 0, (
        "Launcher self-check failed under a clean PATH.\n"
        f"stdout:\n{result.stdout}\n"
        f"stderr:\n{result.stderr}"
    )
    assert "[OK] Launcher self-check passed" in result.stdout
    assert "./.venv/bin/python" in result.stdout, (
        "Launcher must select the project .venv when one exists. "
        "Selecting any other interpreter is what caused the original incident."
    )


@pytest.mark.skipif(os.name == "nt", reason="bash launcher is POSIX-only")
def test_test_dashboard_shell_self_check_with_clean_path() -> None:
    if not (PACKAGE_ROOT / ".venv" / "bin" / "python").exists():
        pytest.skip("project .venv not built; create it via `python3.12 -m venv .venv`")
    env = {
        "HOME": os.environ.get("HOME", "/tmp"),
        "PATH": "/usr/bin:/bin:/usr/sbin:/sbin",
        "SHELL": "/bin/bash",
    }
    result = subprocess.run(
        ["bash", str(TEST_DASH_SH), "--self-check"],
        cwd=str(PACKAGE_ROOT),
        env=env,
        capture_output=True,
        text=True,
        timeout=120,
    )
    assert result.returncode == 0, (
        "run_test_dashboard.sh --self-check failed under a clean PATH.\n"
        f"stdout:\n{result.stdout}\nstderr:\n{result.stderr}"
    )
    assert "[OK] test_dashboard launcher self-check passed" in result.stdout


# ---------------------------------------------------------------------------
# 6. Importability sanity check at the test level (cheap regression net for
#    `dataclass(slots=True)` style breakage).
# ---------------------------------------------------------------------------


def test_pricing_ui_imports_under_supported_python() -> None:
    major, minor = _requires_python_from_pyproject()
    assert sys.version_info[:2] >= (major, minor), (
        f"Test interpreter is {sys.version_info[:2]}; project requires >= "
        f"{(major, minor)}. CI matrix should not run unsupported versions."
    )
    import pricing_ui  # noqa: F401
