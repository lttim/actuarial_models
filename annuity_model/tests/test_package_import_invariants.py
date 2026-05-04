"""Package import invariants for the transitional flat-module layout."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[2]
PACKAGE_ROOT = REPO_ROOT / "annuity_model"

pytestmark = [pytest.mark.invariant]


def _run_import_check(cwd: Path) -> subprocess.CompletedProcess[str]:
    code = """
import annuity_model
import pricing_projection as flat_pricing
import product_registry as flat_registry
import annuity_model.pricing_projection as package_pricing
import annuity_model.product_registry as package_registry

assert annuity_model.SPIAContract is flat_pricing.SPIAContract
assert package_pricing is flat_pricing
assert package_registry is flat_registry
assert annuity_model.get_product_adapter(flat_registry.ProductType.SPIA).is_available()
"""
    return subprocess.run(
        [sys.executable, "-c", code],
        cwd=str(cwd),
        capture_output=True,
        text=True,
        timeout=120,
    )


@pytest.mark.parametrize(
    "cwd",
    [
        pytest.param(REPO_ROOT, id="repo-root"),
        pytest.param(PACKAGE_ROOT, id="annuity-model-cwd"),
    ],
)
def test_annuity_model_imports_from_supported_working_directories(cwd: Path) -> None:
    result = _run_import_check(cwd)
    assert result.returncode == 0, result.stderr + result.stdout
