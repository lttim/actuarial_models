"""Runtime dependency manifest invariants.

Streamlit Community Cloud installs the repository-root ``requirements.txt``
before executing ``streamlit_app.py``. CI and local development also use the
locked dependency tree, so the Cloud manifest needs its own guard: direct
runtime imports in app code must be present in the loose runtime manifests.
"""

from __future__ import annotations

import ast
import sys
import tomllib
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[2]
PROJECT_ROOT = REPO_ROOT / "annuity_model"
SRC_ROOT = PROJECT_ROOT / "src" / "annuity_model"

ROOT_REQUIREMENTS = REPO_ROOT / "requirements.txt"
PROJECT_REQUIREMENTS = PROJECT_ROOT / "requirements.txt"
PYPROJECT = PROJECT_ROOT / "pyproject.toml"
LOCKFILE = PROJECT_ROOT / "requirements.lock"

pytestmark = [pytest.mark.invariant]

OPTIONAL_IMPORTS: dict[str, set[str]] = {
    "opentelemetry": {"annuity_model/src/annuity_model/_observability.py"},
}


def _parse_pinned_requirements(path: Path) -> dict[str, str]:
    pins: dict[str, str] = {}
    for line_no, raw in enumerate(path.read_text(encoding="utf-8").splitlines(), start=1):
        stripped = raw.split("#", 1)[0].strip()
        if not stripped:
            continue
        if "==" not in stripped:
            raise AssertionError(f"{path}:{line_no} is not an exact == pin: {raw!r}")
        name, version = stripped.split("==", 1)
        pins[name.lower().replace("_", "-")] = version
    return pins


def _pyproject_runtime_deps() -> dict[str, str]:
    data = tomllib.loads(PYPROJECT.read_text(encoding="utf-8"))
    pins: dict[str, str] = {}
    for raw in data["project"]["dependencies"]:
        if "==" not in raw:
            raise AssertionError(f"pyproject runtime dependency is not an exact == pin: {raw!r}")
        name, version = raw.split("==", 1)
        pins[name.lower().replace("_", "-")] = version
    return pins


def _source_files_for_runtime_scan() -> list[Path]:
    return [REPO_ROOT / "streamlit_app.py", *sorted(SRC_ROOT.rglob("*.py"))]


def _required_import_roots() -> dict[str, set[str]]:
    roots: dict[str, set[str]] = {}
    stdlib = sys.stdlib_module_names
    for path in _source_files_for_runtime_scan():
        repo_path = path.relative_to(REPO_ROOT).as_posix()
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for node in ast.walk(tree):
            candidates: list[str] = []
            if isinstance(node, ast.Import):
                candidates = [alias.name.split(".", 1)[0] for alias in node.names]
            elif isinstance(node, ast.ImportFrom) and node.level == 0 and node.module:
                candidates = [node.module.split(".", 1)[0]]
            for root in candidates:
                if root == "annuity_model" or root in stdlib or root.startswith("_"):
                    continue
                if repo_path in OPTIONAL_IMPORTS.get(root, set()):
                    continue
                roots.setdefault(root.lower().replace("_", "-"), set()).add(
                    f"{repo_path}:{node.lineno}"
                )
    return roots


def test_runtime_manifests_match_exactly() -> None:
    root_reqs = _parse_pinned_requirements(ROOT_REQUIREMENTS)
    project_reqs = _parse_pinned_requirements(PROJECT_REQUIREMENTS)
    pyproject_reqs = _pyproject_runtime_deps()

    assert root_reqs == project_reqs, (
        "Root requirements.txt is the Streamlit Cloud install manifest and "
        "must stay exactly mirrored with annuity_model/requirements.txt."
    )
    assert root_reqs == pyproject_reqs, (
        "Runtime requirements.txt files and pyproject.toml [project].dependencies "
        "must declare the same direct runtime dependencies."
    )


def test_pyproject_runtime_pins_exist_in_lockfile() -> None:
    pyproject_reqs = _pyproject_runtime_deps()
    lock_reqs = _parse_pinned_requirements(LOCKFILE)
    missing = sorted(set(pyproject_reqs) - set(lock_reqs))
    mismatched = sorted(
        name
        for name, version in pyproject_reqs.items()
        if name in lock_reqs and lock_reqs[name] != version
    )
    assert missing == [], f"requirements.lock is missing runtime deps: {missing}"
    assert mismatched == [], "requirements.lock has runtime dependency pin drift: " + ", ".join(
        f"{name} pyproject={pyproject_reqs[name]} lock={lock_reqs[name]}" for name in mismatched
    )


def test_runtime_imports_are_declared_in_cloud_manifest() -> None:
    root_reqs = _parse_pinned_requirements(ROOT_REQUIREMENTS)
    imported = _required_import_roots()
    missing = {name: locations for name, locations in imported.items() if name not in root_reqs}
    assert missing == {}, (
        "Runtime code imports packages missing from root requirements.txt "
        "(the Streamlit Cloud production manifest): "
        + "; ".join(
            f"{name} at {', '.join(sorted(locations)[:5])}"
            for name, locations in sorted(missing.items())
        )
    )
