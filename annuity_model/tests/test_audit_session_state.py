"""Tests for ``scripts/audit_session_state.py``.

The audit script is the prerequisite enabler the ``ui/MIGRATION.md``
plan demands for splitting ``pricing_ui.py`` into per-page modules.
These tests pin its public behaviour so a future refactor of the AST
walker (or of ``pricing_ui.py`` itself) cannot silently degrade the
audit.
"""

from __future__ import annotations

import ast
import json
import sys
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = REPO_ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))

import audit_session_state as audit  # noqa: E402  (sys.path manipulation)


def _audit_source(src: str) -> audit.AuditReport:
    return audit._audit(ast.parse(src))


def test_subscript_literal_recorded() -> None:
    src = """
def _render_demo():
    import streamlit as st
    return st.session_state["pricing_res"]
"""
    rep = _audit_source(src)
    assert rep.literal_count == 1
    assert rep.symbol_count == 0
    assert "pricing_res" in rep.page_to_keys["_render_demo"]


def test_get_method_literal_recorded() -> None:
    src = """
def _render_demo():
    import streamlit as st
    return st.session_state.get("alm_last", None)
"""
    rep = _audit_source(src)
    assert rep.usages[0].via == "method:get"
    assert rep.usages[0].is_literal is True


def test_widget_key_literal_recorded() -> None:
    src = """
def _render_demo():
    import streamlit as st
    st.text_input("Label", key="my_text")
"""
    rep = _audit_source(src)
    assert rep.usages[0].via == "widget_key"
    assert "my_text" in rep.page_to_keys["_render_demo"]


def test_run_key_symbol_recorded_as_non_literal() -> None:
    """A reference via ``RUN_KEY.X`` is the *good* path -- it should be
    counted, but flagged as ``is_literal=False`` so the literal vs
    symbol breakdown reflects migration progress."""
    src = """
import streamlit as st
from pricing_run_form_state import RUN_KEY
def _render_demo():
    return st.session_state[RUN_KEY.ISSUE_AGE]
"""
    rep = _audit_source(src)
    assert rep.symbol_count == 1
    assert rep.literal_count == 0
    # Use the RUN_KEY symbol rather than the raw issue-age string
    # literal so this test honours the no-new-raw-literals ratchet
    # enforced by tests/test_run_state_key_drift.py.
    from pricing_run_form_state import RUN_KEY

    assert RUN_KEY.ISSUE_AGE in rep.page_to_keys["_render_demo"]


def test_cross_page_keys_detected() -> None:
    src = """
def _render_a():
    import streamlit as st
    return st.session_state["shared"]
def _render_b():
    import streamlit as st
    return st.session_state["shared"]
def _render_c():
    import streamlit as st
    return st.session_state["only_c"]
"""
    rep = _audit_source(src)
    cross = rep.cross_page_keys()
    assert "shared" in cross
    assert sorted(cross["shared"]) == ["_render_a", "_render_b"]
    assert "only_c" not in cross


def test_non_render_functions_ignored() -> None:
    """Helpers without ``_render_`` prefix are NOT treated as pages.

    Otherwise small utility functions would wrongly inflate the
    cross-page count.
    """
    src = """
def helper():
    import streamlit as st
    return st.session_state["x"]
def _render_real():
    import streamlit as st
    return st.session_state["y"]
"""
    rep = _audit_source(src)
    assert "helper" not in rep.page_to_keys
    assert "_render_real" in rep.page_to_keys
    assert rep.page_to_keys["_render_real"] == {"y"}


def test_fail_on_cross_page_returns_one_when_unallowed_shared(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    fake_ui = tmp_path / "fake_pricing_ui.py"
    fake_ui.write_text(
        "def _render_a():\n"
        "    import streamlit as st\n"
        "    return st.session_state['k1']\n"
        "def _render_b():\n"
        "    import streamlit as st\n"
        "    return st.session_state['k1']\n"
    )
    monkeypatch.setattr(audit, "PRICING_UI", fake_ui)
    rc = audit.main(["--fail-on-cross-page"])
    assert rc == 1


def test_fail_on_cross_page_returns_zero_when_allowed(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    fake_ui = tmp_path / "fake_pricing_ui.py"
    fake_ui.write_text(
        "def _render_a():\n"
        "    import streamlit as st\n"
        "    return st.session_state['allowed']\n"
        "def _render_b():\n"
        "    import streamlit as st\n"
        "    return st.session_state['allowed']\n"
    )
    monkeypatch.setattr(audit, "PRICING_UI", fake_ui)
    rc = audit.main(["--fail-on-cross-page", "--allow-cross-page", "allowed"])
    assert rc == 0


def test_json_output_is_parseable(
    capsys: pytest.CaptureFixture[str],
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    fake_ui = tmp_path / "fake_pricing_ui.py"
    fake_ui.write_text(
        "def _render_a():\n"
        "    import streamlit as st\n"
        "    return st.session_state['x']\n"
    )
    monkeypatch.setattr(audit, "PRICING_UI", fake_ui)
    rc = audit.main(["--json"])
    assert rc == 0
    payload = json.loads(capsys.readouterr().out)
    assert payload["totals"]["unique_keys"] == 1
    assert payload["per_page"]["_render_a"] == ["x"]


def test_real_pricing_ui_audits_without_error() -> None:
    """End-to-end smoke: the script must run cleanly against the real
    ``pricing_ui.py`` and produce a non-zero per-page key count for
    each ``_render_<page>`` function it finds.
    """
    rep = audit._audit(ast.parse(audit.PRICING_UI.read_text()))
    assert len(rep.page_to_keys) >= 4, (
        "audit found <4 pages -- pricing_ui.py renamed a _render_ "
        "function? update audit_session_state.py page detection."
    )
    assert all(len(keys) > 0 for keys in rep.page_to_keys.values())
