"""Tests for ``test_dashboard`` helpers (no Streamlit runtime required for most cases)."""

from __future__ import annotations

from pathlib import Path
from textwrap import dedent

import pytest

import test_dashboard as td


def test_discover_tests_metadata_non_empty() -> None:
    rows = td.discover_tests_metadata()
    assert rows, "expected tests/test_pricing_projection.py to define tests"
    names = {r["name"] for r in rows}
    assert "test_yield_curve_from_flat_rate_discount_factors" in names
    for r in rows:
        assert r["section"]
        assert "description" in r


def test_parse_junit_results_empty_when_missing_file(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr(td, "JUNIT_PATH", tmp_path / "nope.xml")
    assert td.parse_junit_results() == {}


def test_parse_junit_results_parametrize_aggregation(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    xml = dedent(
        """\
        <?xml version="1.0" encoding="utf-8"?>
        <testsuites>
          <testsuite name="ts" tests="2" failures="1" errors="0" skipped="0">
            <testcase classname="c" name="test_foo[a]" time="0.1">
              <failure message="bad">trace</failure>
            </testcase>
            <testcase classname="c" name="test_foo[b]" time="0.2"/>
          </testsuite>
        </testsuites>
        """
    )
    path = tmp_path / "junit.xml"
    path.write_text(xml, encoding="utf-8")
    monkeypatch.setattr(td, "JUNIT_PATH", path)
    out = td.parse_junit_results()

    assert out["test_foo"]["status"] == "failed"
    assert "bad" in out["test_foo"]["message"]


def test_run_pytest_junit_returns_code_two_without_interpreter(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(td, "select_pytest_interpreter", lambda root: (None, "no interpreter"))
    code, tail = td.run_pytest_junit()
    assert code == 2
    assert "no interpreter" in tail
