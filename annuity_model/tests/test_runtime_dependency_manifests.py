"""Runtime dependency manifest invariants.

Streamlit Community Cloud installs the repository-root ``requirements.txt``
before executing ``streamlit_app.py``. That root manifest is the Cloud app
surface: product runtime dependencies plus the minimal test dependencies
needed by the online Unit Tests tab. The product manifest and pyproject runtime
dependencies stay runtime-only.
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
CLOUD_TEST_EXTRAS = {"pytest", "hypothesis"}

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


def _pyproject_dev_deps() -> dict[str, str]:
    data = tomllib.loads(PYPROJECT.read_text(encoding="utf-8"))
    pins: dict[str, str] = {}
    for raw in data["project"]["optional-dependencies"]["dev"]:
        if "==" not in raw:
            raise AssertionError(f"pyproject dev dependency is not an exact == pin: {raw!r}")
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

    assert project_reqs == pyproject_reqs, (
        "annuity_model/requirements.txt and pyproject.toml [project].dependencies "
        "must declare the same direct product runtime dependencies."
    )

    root_extras = set(root_reqs) - set(project_reqs)
    missing_runtime = set(project_reqs) - set(root_reqs)
    assert missing_runtime == set(), (
        "Root requirements.txt is the Streamlit Cloud app manifest and must include every "
        f"product runtime dependency: {sorted(missing_runtime)}"
    )
    assert root_extras == CLOUD_TEST_EXTRAS, (
        "Root requirements.txt may differ from annuity_model/requirements.txt only by the "
        f"minimal Cloud Unit Tests extras {sorted(CLOUD_TEST_EXTRAS)}; saw {sorted(root_extras)}."
    )
    for name, version in project_reqs.items():
        assert root_reqs[name] == version, (
            f"Root requirements runtime pin drift for {name}: "
            f"root={root_reqs[name]} project={version}"
        )


def test_cloud_unit_test_dependency_pins_match_dev_and_lockfile() -> None:
    root_reqs = _parse_pinned_requirements(ROOT_REQUIREMENTS)
    dev_reqs = _parse_pinned_requirements(PROJECT_ROOT / "requirements-dev.txt")
    pyproject_dev_reqs = _pyproject_dev_deps()
    lock_reqs = _parse_pinned_requirements(LOCKFILE)

    for name in sorted(CLOUD_TEST_EXTRAS):
        assert name in root_reqs, f"root requirements.txt missing Cloud Unit Tests dep {name}"
        assert root_reqs[name] == dev_reqs.get(name), (
            f"Cloud Unit Tests pin drift for {name}: "
            f"root={root_reqs[name]} requirements-dev={dev_reqs.get(name)}"
        )
        assert root_reqs[name] == pyproject_dev_reqs.get(name), (
            f"Cloud Unit Tests pin drift for {name}: "
            f"root={root_reqs[name]} pyproject dev={pyproject_dev_reqs.get(name)}"
        )
        assert root_reqs[name] == lock_reqs.get(name), (
            f"Cloud Unit Tests pin drift for {name}: "
            f"root={root_reqs[name]} lock={lock_reqs.get(name)}"
        )

    unexpected_root_test_deps = {
        name for name in root_reqs if name in pyproject_dev_reqs and name not in CLOUD_TEST_EXTRAS
    }
    assert unexpected_root_test_deps == set(), (
        "Root requirements.txt should not pull the full dev toolchain into Streamlit Cloud: "
        + ", ".join(sorted(unexpected_root_test_deps))
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
