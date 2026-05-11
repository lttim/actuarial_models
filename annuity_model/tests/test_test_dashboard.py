"""Tests for ``test_dashboard`` helpers (no Streamlit runtime required for most cases)."""

from __future__ import annotations

from pathlib import Path
from textwrap import dedent

import pytest

from annuity_model import test_dashboard as td


class _FakeExpander:
    def __init__(self, owner: _FakeStreamlit, label: str) -> None:
        self.owner = owner
        self.label = label

    def __enter__(self) -> _FakeExpander:
        self.owner.expanders.append(self.label)
        return self

    def __exit__(self, *exc_info: object) -> bool:
        return False


class _FakeStreamlit:
    def __init__(self) -> None:
        self.session_state: dict[str, object] = {}
        self.captions: list[str] = []
        self.codes: list[str] = []
        self.errors: list[str] = []
        self.expanders: list[str] = []
        self.markdowns: list[str] = []
        self.subheaders: list[str] = []
        self.warnings: list[str] = []

    def caption(self, text: str) -> None:
        self.captions.append(text)

    def code(self, text: str, language: str | None = None) -> None:
        self.codes.append(text)

    def error(self, text: str) -> None:
        self.errors.append(text)

    def expander(self, label: str, expanded: bool = False) -> _FakeExpander:
        return _FakeExpander(self, label)

    def markdown(self, text: str) -> None:
        self.markdowns.append(text)

    def subheader(self, text: str) -> None:
        self.subheaders.append(text)

    def warning(self, text: str) -> None:
        self.warnings.append(text)


def test_discover_tests_metadata_non_empty() -> None:
    rows = td.discover_tests_metadata()
    assert len(rows) > 500, "expected dashboard discovery to reflect the full pytest suite"
    nodeids = {r["nodeid"] for r in rows}
    assert (
        "tests/test_pricing_projection.py::test_yield_curve_from_flat_rate_discount_factors"
        in nodeids
    )
    assert any(
        n.startswith("tests/test_regression_matrix.py::test_regression_matrix_cell[")
        for n in nodeids
    )
    for r in rows:
        assert r["nodeid"]
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
            <testcase classname="tests.test_example" name="test_foo[a]" time="0.1">
              <failure message="bad">trace</failure>
            </testcase>
            <testcase classname="tests.test_example" name="test_foo[b]" time="0.2"/>
          </testsuite>
        </testsuites>
        """
    )
    path = tmp_path / "junit.xml"
    path.write_text(xml, encoding="utf-8")
    monkeypatch.setattr(td, "JUNIT_PATH", path)
    out = td.parse_junit_results()

    assert out["tests/test_example.py::test_foo[a]"]["status"] == "failed"
    assert out["tests/test_example.py::test_foo[b]"]["status"] == "passed"
    assert "bad" in out["tests/test_example.py::test_foo[a]"]["message"]


def test_run_pytest_junit_returns_code_two_without_interpreter(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(td, "select_pytest_interpreter", lambda root: (None, "no interpreter"))
    code, tail = td.run_pytest_junit()
    assert code == 2
    assert "no interpreter" in tail


def test_discover_tests_metadata_with_error_preserves_collection_diagnostics(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(td, "_collect_nodeids", lambda pytest_args=None: ([], "collect failed"))

    rows, err = td.discover_tests_metadata_with_error()

    assert rows == []
    assert err == "collect failed"
    assert td.discover_tests_metadata() == []


def test_render_unit_tests_page_missing_pytest_shows_dev_only_state(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    fake_st = _FakeStreamlit()
    monkeypatch.setattr(td, "st", fake_st)
    monkeypatch.setattr(td, "select_pytest_interpreter", lambda root: (None, "pytest missing"))

    td.render_unit_tests_page(embedded=True)

    assert fake_st.subheaders == ["Unit tests"]
    assert any("local development tooling" in message for message in fake_st.warnings)
    assert any("requirements-dev.txt" in block for block in fake_st.codes)
    assert any("pytest missing" in block for block in fake_st.codes)
    assert not fake_st.errors
    assert all("Could not collect tests" not in text for text in fake_st.errors + fake_st.warnings)


def test_render_unit_tests_page_surfaces_collection_failure(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    fake_st = _FakeStreamlit()
    monkeypatch.setattr(td, "st", fake_st)
    monkeypatch.setattr(td, "select_pytest_interpreter", lambda root: ("/tmp/python", None))
    monkeypatch.setattr(
        td,
        "discover_tests_metadata_with_error",
        lambda pytest_args=None: ([], "pytest collect traceback"),
    )

    td.render_unit_tests_page(embedded=True)

    assert fake_st.errors == ["Pytest collection did not return any tests."]
    assert fake_st.expanders == ["pytest collection output"]
    assert any("pytest collect traceback" in block for block in fake_st.codes)
