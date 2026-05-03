"""Run the post-stall deep assessment gate bundle.

This script is intentionally a thin orchestrator over the repo's canonical
commands. It does not replace CI; it gives maintainers and agents one local
entry point for the Excel/parity confidence sweep after workbook-generation
changes.
"""

from __future__ import annotations

import argparse
import subprocess
import sys
import time
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class Step:
    name: str
    command: tuple[str, ...]
    cwd: Path


def _repo_root() -> Path:
    return Path(__file__).resolve().parents[2]


def _run_step(step: Step) -> None:
    print(f"\n[deep-assessment] {step.name}")
    print(f"[deep-assessment] cwd={step.cwd}")
    print("[deep-assessment] $ " + " ".join(step.command))
    started = time.monotonic()
    proc = subprocess.run(step.command, cwd=step.cwd, check=False)
    elapsed = time.monotonic() - started
    if proc.returncode != 0:
        raise SystemExit(
            f"[deep-assessment] FAILED: {step.name} exited {proc.returncode} after {elapsed:.1f}s"
        )
    print(f"[deep-assessment] OK: {step.name} ({elapsed:.1f}s)")


def _steps(*, include_portfolio: bool, skip_pre_commit: bool) -> list[Step]:
    root = _repo_root()
    product = root / "annuity_model"
    py = sys.executable

    steps: list[Step] = []
    if not skip_pre_commit:
        steps.append(
            Step(
                "pre-commit (all files)",
                (py, "-m", "pre_commit", "run", "--all-files", "--show-diff-on-failure"),
                root,
            )
        )
    steps.extend(
        [
            Step("parity gate", (py, "-m", "pytest", "tests/parity", "-q"), product),
            Step("full pytest", (py, "-m", "pytest", "-q"), product),
            Step("deep smoke", (py, "scripts/deep_smoke.py"), product),
            Step(
                "parity contract render check",
                (py, "scripts/render_parity_contract.py", "--check"),
                product,
            ),
            Step("documentation map check", (py, "scripts/check_documentation_map.py"), product),
        ]
    )
    if include_portfolio:
        steps.append(Step("portfolio acceptance", ("just", "portfolio-acceptance"), root))
    return steps


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--include-portfolio",
        action="store_true",
        help="Also run the optional portfolio acceptance superset.",
    )
    parser.add_argument(
        "--skip-pre-commit",
        action="store_true",
        help="Skip pre-commit when diagnosing only runtime/parity gates.",
    )
    args = parser.parse_args(argv)

    started = time.monotonic()
    for step in _steps(
        include_portfolio=bool(args.include_portfolio),
        skip_pre_commit=bool(args.skip_pre_commit),
    ):
        _run_step(step)
    elapsed = time.monotonic() - started
    print(f"\n[deep-assessment] PASS: all requested gates passed in {elapsed:.1f}s")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
