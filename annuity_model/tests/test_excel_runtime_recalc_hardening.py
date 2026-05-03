from __future__ import annotations

import multiprocessing as mp
import time
from pathlib import Path

import pytest

import excel_runtime_recalc as recalc


def _hold_recalc_lock(flag: mp.Event, sleep_s: float) -> None:
    with recalc._libreoffice_recalc_lock(timeout=5.0):  # noqa: SLF001
        flag.set()
        time.sleep(float(sleep_s))


def test_soffice_command_uses_isolated_user_installation(tmp_path: Path) -> None:
    cmd = recalc._soffice_command(  # noqa: SLF001 - deliberate harness regression guard
        soffice="/opt/homebrew/bin/soffice",
        profile_dir=tmp_path / "profile",
        out_dir=tmp_path / "out",
        in_path=tmp_path / "in.xlsx",
    )
    assert cmd[0] == "/opt/homebrew/bin/soffice"
    assert cmd[1].startswith("-env:UserInstallation=file://")
    assert "--headless" in cmd
    assert "--convert-to" in cmd


def test_soffice_env_forces_headless_profile_dirs(tmp_path: Path) -> None:
    env = recalc._soffice_env(tmp_path)  # noqa: SLF001 - deliberate harness regression guard
    assert env["HOME"] == str(tmp_path)
    assert env["TMPDIR"] == str(tmp_path)
    assert env["SAL_USE_VCLPLUGIN"]
    assert "headless" in env["JAVA_TOOL_OPTIONS"]


def test_macos_runtime_recalc_is_opt_in(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(recalc.platform, "system", lambda: "Darwin")
    monkeypatch.delenv(recalc.MACOS_LIBREOFFICE_RECALC_ENV, raising=False)
    reason = recalc.libreoffice_recalc_disabled_reason()
    assert reason is not None
    assert recalc.MACOS_LIBREOFFICE_RECALC_ENV in reason
    monkeypatch.setenv(recalc.MACOS_LIBREOFFICE_RECALC_ENV, "1")
    assert recalc.libreoffice_recalc_disabled_reason() is None


def test_libreoffice_recalc_lock_serializes_processes() -> None:
    entered = mp.Event()
    proc = mp.Process(target=_hold_recalc_lock, args=(entered, 0.25))
    proc.start()
    assert entered.wait(timeout=2.0)
    start = time.monotonic()
    with recalc._libreoffice_recalc_lock(timeout=5.0):  # noqa: SLF001
        elapsed = time.monotonic() - start
    proc.join(timeout=2.0)
    assert proc.exitcode == 0
    assert elapsed >= 0.15


def test_libreoffice_recalc_lock_times_out_when_held() -> None:
    entered = mp.Event()
    proc = mp.Process(target=_hold_recalc_lock, args=(entered, 1.5))
    proc.start()
    assert entered.wait(timeout=2.0)
    try:
        with pytest.raises(recalc.RecalcTimeout, match="LibreOffice recalc lock"):
            with recalc._libreoffice_recalc_lock(timeout=0.1):  # noqa: SLF001
                pass
    finally:
        proc.join(timeout=2.0)
        if proc.is_alive():
            proc.terminate()
            proc.join(timeout=2.0)
