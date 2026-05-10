"""Package import invariants for the standard ``src/`` layout."""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[2]
PROJECT_ROOT = REPO_ROOT / "annuity_model"
SRC_ROOT = PROJECT_ROOT / "src"

pytestmark = [pytest.mark.invariant]


def _run_import_check(cwd: Path) -> subprocess.CompletedProcess[str]:
    code = """
import pathlib
import annuity_model
import annuity_model.pricing_projection as pricing_projection
import annuity_model.product_registry as product_registry

package_path = pathlib.Path(annuity_model.__file__).resolve()
assert "/src/annuity_model/" in package_path.as_posix(), package_path
assert annuity_model.SPIAContract is pricing_projection.SPIAContract
assert annuity_model.get_product_adapter(product_registry.ProductType.SPIA).is_available()

try:
    import pricing_projection as bare_pricing_projection
except ModuleNotFoundError:
    bare_pricing_projection = None
assert bare_pricing_projection is None, "bare flat imports must not resolve in src layout"
"""
    env = os.environ.copy()
    env["PYTHONPATH"] = str(SRC_ROOT)
    return subprocess.run(
        [sys.executable, "-c", code],
        cwd=str(cwd),
        env=env,
        capture_output=True,
        text=True,
        timeout=120,
    )


@pytest.mark.parametrize(
    "cwd",
    [
        pytest.param(REPO_ROOT, id="repo-root"),
        pytest.param(PROJECT_ROOT, id="annuity-model-cwd"),
    ],
)
def test_annuity_model_imports_from_supported_working_directories(cwd: Path) -> None:
    result = _run_import_check(cwd)
    assert result.returncode == 0, result.stderr + result.stdout
