"""
Browser dashboard for SPIA unit tests: descriptions, run controls, and outcomes.

Run from the annuity_model folder:
    streamlit run test_dashboard.py
Or double-click run_test_dashboard.bat (Windows).

For the full model workspace (pricing + charts + these tests), use:
    streamlit run spia_ui.py
Or run_spia_ui.bat.
"""

from __future__ import annotations

import ast
import subprocess
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any

import streamlit as st

ROOT = Path(__file__).resolve().parent
TEST_FILE = ROOT / "tests" / "test_spia_projection.py"
REPORTS_DIR = ROOT / "reports"
JUNIT_PATH = REPORTS_DIR / "junit.xml"


def _section_at_line(lines: list[str], lineno: int) -> str:
    section = "General"
    for i in range(min(max(lineno - 1, 0), len(lines))):
        s = lines[i].strip()
        if s.startswith("# ---") and s.endswith("---"):
            section = s[4:-3].strip()
    return section


def discover_tests_metadata() -> list[dict[str, Any]]:
    """Parse test file: name, section (from # --- headers), description (docstring)."""
    if not TEST_FILE.is_file():
        return []
    src = TEST_FILE.read_text(encoding="utf-8")
    lines = src.splitlines()
    tree = ast.parse(src)
    rows: list[dict[str, Any]] = []
    for node in tree.body:
        if isinstance(node, ast.FunctionDef) and node.name.startswith("test_"):
            doc = ast.get_docstring(node)
            desc = doc.strip() if doc else "_(No docstring — add one under the def line in the test file.)_"
            rows.append(
                {
                    "name": node.name,
                    "section": _section_at_line(lines, node.lineno),
                    "description": desc,
                }
            )
    return rows


def run_pytest_junit() -> tuple[int, str]:
    """Run pytest; write JUnit XML. Returns (exit_code, stderr+stdout snippet)."""
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    cmd = [
        sys.executable,
        "-m",
        "pytest",
        str(ROOT / "tests"),
        "-v",
        "--tb=short",
        f"--junitxml={JUNIT_PATH}",
        "-o",
        "junit_family=xunit2",
    ]
    proc = subprocess.run(
        cmd,
        cwd=str(ROOT),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    tail = (proc.stdout or "")[-8000:] + "\n" + (proc.stderr or "")[-4000:]
    return proc.returncode, tail


_STATUS_SEVERITY = {"error": 4, "failed": 3, "skipped": 2, "passed": 1, "not_run": 0}


def _base_test_name(name: str) -> str:
    """Strip parametrize suffix '[...]' so parametrized cases map to their base function name."""
    bracket = name.find("[")
    return name[:bracket] if bracket != -1 else name


def parse_junit_results() -> dict[str, dict[str, Any]]:
    """
    Map test function name -> {status, message, time_s}.
    status in passed | failed | skipped | error | unknown.

    Parametrized tests (e.g. test_foo[case0]) are aggregated under their base
    function name (test_foo) so the dashboard can match them to AST-discovered
    test names.  The most severe status across all cases wins; failure messages
    from all failing/erroring cases are concatenated.
    """
    if not JUNIT_PATH.is_file():
        return {}
    try:
        tree = ET.parse(JUNIT_PATH)
    except ET.ParseError:
        return {}
    root = tree.getroot()

    raw: dict[str, list[dict[str, Any]]] = {}
    for tc in root.iter("testcase"):
        full_name = tc.get("name") or ""
        base = _base_test_name(full_name)
        time_s = tc.get("time")
        fail = tc.find("failure")
        skip = tc.find("skipped")
        err = tc.find("error")
        if fail is not None:
            msg = fail.get("message") or ""
            text = (fail.text or "").strip()
            detail = (msg + "\n" + text).strip() or "Failed"
            entry = {"status": "failed", "message": detail, "time_s": time_s}
        elif err is not None:
            entry = {
                "status": "error",
                "message": (err.get("message") or err.text or "Error").strip(),
                "time_s": time_s,
            }
        elif skip is not None:
            entry = {
                "status": "skipped",
                "message": (skip.get("message") or skip.text or "Skipped").strip(),
                "time_s": time_s,
            }
        else:
            entry = {"status": "passed", "message": "", "time_s": time_s}
        raw.setdefault(base, []).append(entry)

    out: dict[str, dict[str, Any]] = {}
    for base, entries in raw.items():
        worst = max(entries, key=lambda e: _STATUS_SEVERITY.get(e["status"], 0))
        messages = [e["message"] for e in entries if e["message"]]
        total_time = None
        try:
            total_time = str(round(sum(float(e["time_s"]) for e in entries if e["time_s"] is not None), 4))
        except (TypeError, ValueError):
            pass
        out[base] = {
            "status": worst["status"],
            "message": "\n\n".join(messages),
            "time_s": total_time,
        }
    return out


def render_unit_tests_page(*, embedded: bool = False) -> None:
    """
    Render the pytest discovery / run / results UI.

    When embedded=True, sidebar controls are inlined so a parent app can own the sidebar.
    """
    if not embedded:
        st.title("SPIA unit test dashboard")
        st.caption(
            "Each row is one automated check of `spia_projection.py`. "
            "Descriptions are taken from the test’s docstring in `tests/test_spia_projection.py`."
        )

    meta = discover_tests_metadata()
    if not meta:
        st.error(f"Could not find tests at `{TEST_FILE}`. Open the `annuity_model` folder as project root.")
        return

    notify = st.session_state.get("last_notify")
    if notify == "pass":
        st.success("Last test run finished with pytest exit code 0 (all executed tests passed).")
    elif notify == "fail":
        st.warning("Last test run reported failures, errors, or a non-zero pytest exit code. See expanders below.")

    def _run_clicked() -> None:
        with st.spinner("Running pytest…"):
            code, log_tail = run_pytest_junit()
        st.session_state["last_exit_code"] = code
        st.session_state["last_log_tail"] = log_tail
        st.session_state["last_results"] = parse_junit_results()
        st.session_state["last_notify"] = "pass" if code == 0 else "fail"

    if embedded:
        st.subheader("Unit tests")
        st.caption(
            "Automated checks for `spia_projection.py`. Descriptions come from test docstrings in "
            "`tests/test_spia_projection.py`."
        )
        b1, b2 = st.columns(2)
        with b1:
            if st.button("Run all tests", type="primary", use_container_width=True, key="pytest_run_embedded"):
                _run_clicked()
        with b2:
            st.caption("Setup: `python -m pip install -r requirements-dev.txt`")
    else:
        with st.sidebar:
            st.header("Run")
            if st.button("Run all tests", type="primary", use_container_width=True):
                _run_clicked()
            st.divider()
            st.markdown(
                "**First time setup:** in a terminal here, run  \n`python -m pip install -r requirements-dev.txt`"
            )
            st.markdown("**CLI alternative:** `python -m pytest` or `run_tests_report.bat` for HTML.")

    results: dict[str, dict[str, Any]] = st.session_state.get("last_results") or {}

    # Summary metrics
    names = [m["name"] for m in meta]
    passed = sum(1 for n in names if results.get(n, {}).get("status") == "passed")
    failed = sum(1 for n in names if results.get(n, {}).get("status") in ("failed", "error"))
    skipped = sum(1 for n in names if results.get(n, {}).get("status") == "skipped")
    not_run = sum(1 for n in names if n not in results)

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total tests", len(meta))
    c2.metric("Passed", passed)
    c3.metric("Failed", failed)
    c4.metric("Skipped", skipped)
    c5.metric("Not run yet", not_run)

    st.divider()

    # Group by section
    sections: dict[str, list[dict[str, Any]]] = {}
    for m in meta:
        sections.setdefault(m["section"], []).append(m)

    _MC_SECTION = "Monte Carlo first principles"
    for section in sorted(sections.keys(), key=lambda s: (s != _MC_SECTION, s == "General", s)):
        st.subheader(section)
        for m in sections[section]:
            name = m["name"]
            r = results.get(name, {})
            status = r.get("status", "not_run")
            icon = {"passed": "✅", "failed": "❌", "error": "⚠️", "skipped": "⏭️", "not_run": "⚪"}.get(
                status, "⚪"
            )
            with st.expander(f"{icon} **{name}** — _{status.replace('_', ' ')}_", expanded=(status in ("failed", "error"))):
                st.markdown(m["description"])
                if r.get("time_s") is not None:
                    st.caption(f"Runtime: {r['time_s']} s")
                if status in ("failed", "error") and r.get("message"):
                    st.code(r["message"], language="text")
        st.divider()

    if st.session_state.get("last_log_tail"):
        with st.expander("Last pytest output (tail)"):
            st.code(st.session_state["last_log_tail"], language="text")


def main() -> None:
    st.set_page_config(page_title="SPIA unit tests", layout="wide")
    render_unit_tests_page(embedded=False)


if __name__ == "__main__":
    main()
