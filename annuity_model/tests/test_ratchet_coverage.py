"""Unit tests for ``scripts/ratchet_coverage.py``.

These tests exercise the script's pure logic by stubbing the ``coverage``
subprocess call. They do *not* invoke the real ``coverage`` binary nor
mutate the on-disk ``pyproject.toml`` -- the script is parameterised on a
``--coverage-cmd`` and we redirect ``PYPROJECT_PATH`` via monkeypatch.
"""

from __future__ import annotations

import sys
import textwrap
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = REPO_ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))

import ratchet_coverage  # noqa: E402  (sys.path manipulation)


def _write_pyproject(tmp_path: Path, fail_under: float | str) -> Path:
    """Build a minimal pyproject with the [tool.coverage.report] table."""
    body = textwrap.dedent(
        f"""
        [tool.coverage.report]
        precision = 1
        fail_under = {fail_under}
        """
    ).strip()
    p = tmp_path / "pyproject.toml"
    p.write_text(body)
    return p


@pytest.fixture()
def fake_coverage_cmd(tmp_path: Path) -> list[str]:
    """Return a ``--coverage-cmd`` shim that prints a configurable number.

    The shim is a tiny Python script invoked via ``sys.executable``, which
    matches how a real CI runner would call into a venv-pinned tool.
    """
    shim = tmp_path / "fake_coverage.py"
    shim.write_text(
        textwrap.dedent(
            """
            import os, sys
            pct = os.environ.get("FAKE_COV_PCT", "60.0")
            sys.stdout.write(pct)
            sys.exit(0)
            """
        ).strip()
    )
    return [sys.executable, str(shim)]


def test_pass_when_actual_above_floor(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
    fake_coverage_cmd: list[str],
) -> None:
    monkeypatch.setattr(ratchet_coverage, "PYPROJECT_PATH", _write_pyproject(tmp_path, 55.0))
    monkeypatch.setenv("FAKE_COV_PCT", "60.0")
    rc = ratchet_coverage.main(["--coverage-cmd", *fake_coverage_cmd])
    assert rc == 0
    out = capsys.readouterr().out
    assert "60.0%" in out
    assert "55.0%" in out
    assert "above floor" in out


def test_fail_when_actual_below_floor(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
    fake_coverage_cmd: list[str],
) -> None:
    monkeypatch.setattr(ratchet_coverage, "PYPROJECT_PATH", _write_pyproject(tmp_path, 70.0))
    monkeypatch.setenv("FAKE_COV_PCT", "65.0")
    rc = ratchet_coverage.main(["--coverage-cmd", *fake_coverage_cmd])
    assert rc == 1
    captured = capsys.readouterr()
    assert "FAIL" in captured.err
    assert "65.0%" in captured.out and "70.0%" in captured.out


def test_bump_hint_fires_when_headroom_exceeds_threshold(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
    fake_coverage_cmd: list[str],
) -> None:
    monkeypatch.setattr(ratchet_coverage, "PYPROJECT_PATH", _write_pyproject(tmp_path, 50.0))
    monkeypatch.setenv("FAKE_COV_PCT", "75.0")
    rc = ratchet_coverage.main(["--coverage-cmd", *fake_coverage_cmd, "--bump-hint", "5.0"])
    assert rc == 0
    out = capsys.readouterr().out
    assert "consider running" in out
    assert "ratchet_coverage.py --update" in out


def test_bump_hint_silent_inside_threshold(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
    fake_coverage_cmd: list[str],
) -> None:
    monkeypatch.setattr(ratchet_coverage, "PYPROJECT_PATH", _write_pyproject(tmp_path, 60.0))
    monkeypatch.setenv("FAKE_COV_PCT", "60.5")
    rc = ratchet_coverage.main(["--coverage-cmd", *fake_coverage_cmd, "--bump-hint", "1.0"])
    assert rc == 0
    out = capsys.readouterr().out
    assert "consider running" not in out


def test_update_writes_new_floor(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
    fake_coverage_cmd: list[str],
) -> None:
    pp = _write_pyproject(tmp_path, 55.0)
    monkeypatch.setattr(ratchet_coverage, "PYPROJECT_PATH", pp)
    monkeypatch.setenv("FAKE_COV_PCT", "62.7")
    rc = ratchet_coverage.main(["--update", "--coverage-cmd", *fake_coverage_cmd])
    assert rc == 0
    text = pp.read_text()
    assert "fail_under = 62.7" in text
    assert "fail_under = 55.0" not in text


def test_update_refuses_to_lower_floor(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
    fake_coverage_cmd: list[str],
) -> None:
    pp = _write_pyproject(tmp_path, 70.0)
    monkeypatch.setattr(ratchet_coverage, "PYPROJECT_PATH", pp)
    monkeypatch.setenv("FAKE_COV_PCT", "60.0")
    rc = ratchet_coverage.main(["--update", "--coverage-cmd", *fake_coverage_cmd])
    assert rc == 3
    err = capsys.readouterr().err
    assert "refusing to update" in err
    assert "fail_under = 70.0" in pp.read_text()


def test_update_no_op_when_actual_rounds_to_current_floor(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
    fake_coverage_cmd: list[str],
) -> None:
    pp = _write_pyproject(tmp_path, 60.0)
    monkeypatch.setattr(ratchet_coverage, "PYPROJECT_PATH", pp)
    monkeypatch.setenv("FAKE_COV_PCT", "60.04")
    rc = ratchet_coverage.main(["--update", "--coverage-cmd", *fake_coverage_cmd])
    assert rc == 0
    out = capsys.readouterr().out
    assert "no update needed" in out
    assert "fail_under = 60.0" in pp.read_text()


def test_missing_fail_under_key_exits_2(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
    fake_coverage_cmd: list[str],
) -> None:
    """A pyproject without ``fail_under`` must abort with exit code 2.

    ``_read_floor`` calls ``sys.exit`` directly (not ``return``) so the
    failure surfaces as ``SystemExit``. We intentionally do *not* swallow
    that into ``main()``'s normal return path because the rest of the
    script cannot recover from a missing floor and we want the failure
    mode to be unambiguous in CI output.
    """
    pp = tmp_path / "pyproject.toml"
    pp.write_text("[tool.coverage.report]\nprecision = 1\n")
    monkeypatch.setattr(ratchet_coverage, "PYPROJECT_PATH", pp)
    with pytest.raises(SystemExit) as excinfo:
        ratchet_coverage.main(["--coverage-cmd", *fake_coverage_cmd])
    assert excinfo.value.code == 2
    err = capsys.readouterr().err
    assert "fail_under" in err


def test_pyproject_has_fail_under_in_real_repo() -> None:
    """The on-disk pyproject must keep `fail_under` -- CI depends on it."""
    import tomllib

    real_pp = REPO_ROOT / "pyproject.toml"
    with real_pp.open("rb") as fh:
        data = tomllib.load(fh)
    floor = data["tool"]["coverage"]["report"]["fail_under"]
    assert isinstance(floor, int | float), (
        "[tool.coverage.report].fail_under must be numeric for "
        "scripts/ratchet_coverage.py to enforce the gate"
    )
    assert 0.0 <= float(floor) <= 100.0
