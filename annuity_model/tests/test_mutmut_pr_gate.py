"""Unit + meta-invariant tests for ``scripts/mutmut_pr_gate.py``.

Coverage strategy
-----------------
* Pure unit tests: stub out ``mutmut run`` (via ``--skip-run``) and feed
  pre-cooked ``mutants/<path>.meta`` JSON files into the parser /
  threshold-comparison logic. This is fast, deterministic, and works on
  every OS (mutmut itself does not).
* Meta-invariant tests: assert that
  ``MUTMUT_SURFACE`` in the gate stays in lockstep with both the
  ``paths_to_mutate`` list embedded in
  ``.github/workflows/mutmut-nightly.yml`` and the per-file thresholds
  declared in ``mutmut_thresholds.toml``. A divergence here is
  almost always a copy-paste-and-forget bug.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = REPO_ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))

import mutmut_pr_gate as gate  # noqa: E402  (sys.path manipulation)


def _seed_meta(mutants_dir: Path, rel_path: str, exit_codes: dict[str, int]) -> None:
    target = mutants_dir / f"{rel_path}.meta"
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps({"exit_code_by_key": exit_codes}, indent=2))


def _write_thresholds(tmp_path: Path, default: int, per_file: dict[str, int]) -> Path:
    body = ["[thresholds]", f"default = {default}", "", "[thresholds.per_file]"]
    for k, v in per_file.items():
        body.append(f'"{k}" = {v}')
    p = tmp_path / "mutmut_thresholds.toml"
    p.write_text("\n".join(body) + "\n")
    return p


def test_no_touched_files_returns_zero(capsys: pytest.CaptureFixture[str]) -> None:
    rc = gate.main(["--skip-run", "--touched-files"])
    assert rc == 0
    assert "no --touched-files" in capsys.readouterr().out


def test_touched_files_outside_surface_returns_zero(
    capsys: pytest.CaptureFixture[str],
) -> None:
    rc = gate.main(["--skip-run", "--touched-files", "tests/test_foo.py", "docs/README.md"])
    assert rc == 0
    assert "no parity-critical files" in capsys.readouterr().out


def test_pass_when_zero_survivors(
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    mutants = tmp_path / "mutants"
    _seed_meta(
        mutants,
        "src/annuity_model/pricing_projection.py",
        {"pp.x__mutmut_1": 1, "pp.x__mutmut_2": 1},  # killed/killed
    )
    thresholds = _write_thresholds(tmp_path, default=0, per_file={})
    rc = gate.main(
        [
            "--skip-run",
            "--touched-files",
            "src/annuity_model/pricing_projection.py",
            "--thresholds",
            str(thresholds),
            "--mutants-dir",
            str(mutants),
        ]
    )
    assert rc == 0
    out = capsys.readouterr().out
    assert "0 survivor(s) / 2 mutants" in out
    assert "PASS" in out


def test_fail_when_survivor_exceeds_default_cap(
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    mutants = tmp_path / "mutants"
    _seed_meta(
        mutants,
        "src/annuity_model/rila_projection.py",
        {"r.x__mutmut_1": 0, "r.x__mutmut_2": 1},  # one survivor
    )
    thresholds = _write_thresholds(tmp_path, default=0, per_file={})
    rc = gate.main(
        [
            "--skip-run",
            "--touched-files",
            "src/annuity_model/rila_projection.py",
            "--thresholds",
            str(thresholds),
            "--mutants-dir",
            str(mutants),
        ]
    )
    assert rc == 1
    captured = capsys.readouterr()
    assert "1 survivor(s) / 2 mutants" in captured.out
    assert "FAIL" in captured.err


def test_per_file_threshold_overrides_default(
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    """A per-file cap of 3 must allow 2 survivors through."""
    mutants = tmp_path / "mutants"
    _seed_meta(
        mutants,
        "src/annuity_model/pricing_projection.py",
        {f"pp.x__mutmut_{i}": (0 if i < 2 else 1) for i in range(5)},
    )
    thresholds = _write_thresholds(
        tmp_path, default=0, per_file={"src/annuity_model/pricing_projection.py": 3}
    )
    rc = gate.main(
        [
            "--skip-run",
            "--touched-files",
            "src/annuity_model/pricing_projection.py",
            "--thresholds",
            str(thresholds),
            "--mutants-dir",
            str(mutants),
        ]
    )
    assert rc == 0
    out = capsys.readouterr().out
    assert "cap = 3" in out


def test_per_file_threshold_can_still_fail_when_exceeded(
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    mutants = tmp_path / "mutants"
    _seed_meta(
        mutants,
        "src/annuity_model/pricing_projection.py",
        {f"pp.x__mutmut_{i}": 0 for i in range(4)},
    )
    thresholds = _write_thresholds(
        tmp_path, default=0, per_file={"src/annuity_model/pricing_projection.py": 3}
    )
    rc = gate.main(
        [
            "--skip-run",
            "--touched-files",
            "src/annuity_model/pricing_projection.py",
            "--thresholds",
            str(thresholds),
            "--mutants-dir",
            str(mutants),
        ]
    )
    assert rc == 1
    out = capsys.readouterr().out
    assert "4 survivor(s)" in out and "cap = 3" in out


def test_missing_meta_file_treated_as_zero_zero(
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    """If mutmut never produced a meta for a touched file, gate must pass."""
    mutants = tmp_path / "mutants"
    mutants.mkdir()
    thresholds = _write_thresholds(tmp_path, default=0, per_file={})
    rc = gate.main(
        [
            "--skip-run",
            "--touched-files",
            "src/annuity_model/pricing_projection.py",
            "--thresholds",
            str(thresholds),
            "--mutants-dir",
            str(mutants),
        ]
    )
    assert rc == 0
    captured = capsys.readouterr()
    assert "no mutants/src/annuity_model/pricing_projection.py.meta produced" in captured.err


def test_negative_threshold_aborts_with_two(
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    """Negative thresholds are nonsense and must be rejected."""
    p = tmp_path / "mutmut_thresholds.toml"
    p.write_text("[thresholds]\ndefault = -1\n")
    with pytest.raises(SystemExit) as excinfo:
        gate.main(["--skip-run", "--thresholds", str(p)])
    assert excinfo.value.code == 2


def test_missing_thresholds_file_aborts_with_two(
    tmp_path: Path,
) -> None:
    missing = tmp_path / "no_such.toml"
    with pytest.raises(SystemExit) as excinfo:
        gate.main(["--skip-run", "--thresholds", str(missing)])
    assert excinfo.value.code == 2


def test_surface_matches_nightly_workflow() -> None:
    """``MUTMUT_SURFACE`` must equal the nightly workflow's
    ``paths_to_mutate`` list, line-for-line.

    This is the meta-invariant that prevents the PR-level gate from
    drifting out of sync with the nightly. If a new engine is added to
    the nightly surface, this test fails until the PR-level gate is
    updated to match (and vice-versa).
    """
    workflow = REPO_ROOT.parent / ".github" / "workflows" / "mutmut-nightly.yml"
    assert workflow.is_file(), f"missing nightly workflow at {workflow}"
    text = workflow.read_text()
    # Extract the inline `paths_to_mutate = [...]` block; it lives inside
    # a heredoc so we can't import it. A simple line scan is plenty.
    in_block = False
    declared: list[str] = []
    for line in text.splitlines():
        stripped = line.strip()
        if stripped.startswith("paths_to_mutate"):
            in_block = True
            continue
        if in_block:
            if stripped == "]":
                break
            if stripped.startswith('"') and stripped.endswith(",") or stripped.endswith('"'):
                declared.append(stripped.strip(",").strip('"'))
    nightly = frozenset(declared)
    assert nightly == gate.MUTMUT_SURFACE, (
        "MUTMUT_SURFACE drift between PR gate and nightly:\n"
        f"  PR gate only:  {sorted(gate.MUTMUT_SURFACE - nightly)}\n"
        f"  nightly only:  {sorted(nightly - gate.MUTMUT_SURFACE)}"
    )


def test_threshold_file_keys_resolve_to_real_surface_files() -> None:
    """Every per-file threshold key must reference a file in
    ``MUTMUT_SURFACE`` AND exist on disk. A typo'd path silently makes
    the override a no-op (the default fires instead), which would mask a
    regression on whichever file the author actually meant.
    """
    _, per_file = gate._load_thresholds()
    bad_keys = [k for k in per_file if k not in gate.MUTMUT_SURFACE]
    assert not bad_keys, (
        "mutmut_thresholds.toml [thresholds.per_file] entries do not match "
        f"MUTMUT_SURFACE: {bad_keys}"
    )
    missing_files = [k for k in per_file if not (REPO_ROOT / k).is_file()]
    assert not missing_files, (
        "mutmut_thresholds.toml [thresholds.per_file] entries reference "
        f"non-existent files: {missing_files}"
    )


def test_every_surface_file_exists_on_disk() -> None:
    """Catches the inverse drift: a surface file gets renamed/deleted
    without updating the gate's allow-list."""
    missing = sorted(p for p in gate.MUTMUT_SURFACE if not (REPO_ROOT / p).is_file())
    assert not missing, (
        "MUTMUT_SURFACE references files that no longer exist on disk: "
        f"{missing}. Update scripts/mutmut_pr_gate.py and the nightly "
        "workflow in the same commit as the rename/delete."
    )


def test_thresholds_real_file_loads() -> None:
    """The on-disk ``mutmut_thresholds.toml`` must parse + validate.

    This tests the real config (no monkeypatch) so a malformed entry
    fails locally before reaching CI.
    """
    default, per_file = gate._load_thresholds()
    assert isinstance(default, int) and default >= 0
    assert isinstance(per_file, dict)
    for v in per_file.values():
        assert isinstance(v, int) and v >= 0
