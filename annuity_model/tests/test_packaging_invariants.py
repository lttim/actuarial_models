"""Build/install invariants for the standard package layout."""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent

pytestmark = [pytest.mark.invariant]


def test_package_builds_and_installs_from_wheel(tmp_path: Path) -> None:
    wheels = tmp_path / "wheels"
    wheels.mkdir()
    shutil.rmtree(PROJECT_ROOT / "build", ignore_errors=True)

    try:
        build = subprocess.run(
            [
                sys.executable,
                "-m",
                "pip",
                "wheel",
                "--no-deps",
                "--no-build-isolation",
                "--wheel-dir",
                str(wheels),
                ".",
            ],
            cwd=PROJECT_ROOT,
            capture_output=True,
            text=True,
            timeout=180,
        )
    finally:
        shutil.rmtree(PROJECT_ROOT / "build", ignore_errors=True)
    assert build.returncode == 0, build.stderr + build.stdout
    wheel = next(wheels.glob("annuity_model-*.whl"))

    target = tmp_path / "site"

    install = subprocess.run(
        [sys.executable, "-m", "pip", "install", "--no-deps", "--target", str(target), str(wheel)],
        capture_output=True,
        text=True,
        timeout=180,
    )
    assert install.returncode == 0, install.stderr + install.stdout

    dist_info = next(target.glob("annuity_model-*.dist-info"))
    env = os.environ.copy()
    env["PYTHONPATH"] = str(target)
    smoke = subprocess.run(
        [
            sys.executable,
            "-c",
            (
                "import importlib.metadata as md; "
                "dist = md.Distribution.at(" + repr(str(dist_info)) + "); "
                "eps = {ep.name for ep in dist.entry_points}; "
                "import annuity_model; "
                "assert 'annuity-portfolio' in eps; "
                "assert 'annuity-pricing-ui' in eps; "
                "assert 'annuity-test-dashboard' in eps; "
                "assert annuity_model.SPIAContract.__module__ == 'annuity_model.pricing_projection'"
            ),
        ],
        cwd=tmp_path,
        env=env,
        capture_output=True,
        text=True,
        timeout=120,
    )
    assert smoke.returncode == 0, smoke.stderr + smoke.stdout

    help_run = subprocess.run(
        [
            sys.executable,
            "-c",
            "from annuity_model.cli import main; raise SystemExit(main(['--help']))",
        ],
        cwd=tmp_path,
        env=env,
        capture_output=True,
        text=True,
        timeout=120,
    )
    assert help_run.returncode == 0, help_run.stderr + help_run.stdout
    assert "portfolio-run" in help_run.stdout
