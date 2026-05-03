"""Tests for ``portfolio_config`` enablement (Streamlit + CLI gate)."""

from __future__ import annotations

from pathlib import Path

import pytest

import portfolio_config as pc
from portfolio_config import portfolio_sidebar_visible, portfolio_v1_enabled
from pricing_run_form_state import PORTFOLIO_KEY


def test_default_on_when_env_unset_and_no_disable_file(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    monkeypatch.setattr(pc, "_repo_root", lambda: tmp_path)
    monkeypatch.delenv("ANNUITY_MODEL_PORTFOLIO_V1", raising=False)
    assert portfolio_v1_enabled() is True


def test_disable_file_turns_off_even_when_env_truthy(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    monkeypatch.setattr(pc, "_repo_root", lambda: tmp_path)
    (tmp_path / ".disable-portfolio-v1").write_text("", encoding="utf-8")
    monkeypatch.setenv("ANNUITY_MODEL_PORTFOLIO_V1", "1")
    assert portfolio_v1_enabled() is False


def test_explicit_zero_disables(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.setattr(pc, "_repo_root", lambda: tmp_path)
    monkeypatch.setenv("ANNUITY_MODEL_PORTFOLIO_V1", "0")
    assert portfolio_v1_enabled() is False


def test_truthy_env_enables_without_disable_file(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    monkeypatch.setattr(pc, "_repo_root", lambda: tmp_path)
    monkeypatch.setenv("ANNUITY_MODEL_PORTFOLIO_V1", "true")
    assert portfolio_v1_enabled() is True


def test_sidebar_visible_force_session(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.setattr(pc, "_repo_root", lambda: tmp_path)
    monkeypatch.setenv("ANNUITY_MODEL_PORTFOLIO_V1", "0")
    assert portfolio_v1_enabled() is False
    assert portfolio_sidebar_visible({PORTFOLIO_KEY.UI_FORCE_SIDEBAR: True}) is True
    assert portfolio_sidebar_visible({PORTFOLIO_KEY.UI_FORCE_SIDEBAR: False}) is False
