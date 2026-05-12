"""
Browser dashboard for pricing engine unit tests: descriptions, run controls, and outcomes.

Run from the annuity_model folder:
    streamlit run src/annuity_model/test_dashboard.py
Or double-click run_test_dashboard.bat (Windows) / ./run_test_dashboard.sh (macOS, Linux).

For the full model workspace (pricing + charts + these tests), use:
    streamlit run src/annuity_model/pricing_ui.py
Or run_pricing_ui.bat / ./run_pricing_ui.sh.
"""

from __future__ import annotations

import ast
import re

# Reviewed: dashboard runs fixed local pytest commands only.
import subprocess  # nosec B404
from pathlib import Path
from typing import Any

import streamlit as st
from defusedxml import ElementTree

from annuity_model.pytest_python import select_pytest_interpreter

ROOT = Path(__file__).resolve().parents[2]
REPORTS_DIR = ROOT / "reports"
JUNIT_PATH = REPORTS_DIR / "junit.xml"
COLLECT_TAIL_RE = re.compile(r"^\d+ tests? collected")
PRESET_ARGS: dict[str, list[str]] = {
    "Full suite": [str(ROOT / "tests")],
    "Parity": [str(ROOT / "tests" / "parity")],
    "UI": [str(ROOT / "tests" / "ui")],
    "Integration": [str(ROOT / "tests" / "integration")],
    "Unit folder": [str(ROOT / "tests" / "unit")],
}
LOCAL_PYTEST_SETUP = """cd annuity_model
python3.12 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt -r requirements-dev.txt
./run_pricing_ui.sh"""


def _section_at_line(lines: list[str], lineno: int) -> str:
    section = "General"
    for i in range(min(max(lineno - 1, 0), len(lines))):
        s = lines[i].strip()
        if s.startswith("# ---") and s.endswith("---"):
            section = s[4:-3].strip()
    return section


def _docstring_index() -> dict[tuple[str, str], dict[str, str]]:
    """Return docstring/section metadata indexed by ``(relpath, test_function)``."""
    out: dict[tuple[str, str], dict[str, str]] = {}
    for path in sorted((ROOT / "tests").rglob("test_*.py")):
        try:
            src = path.read_text(encoding="utf-8")
            tree = ast.parse(src)
        except (OSError, SyntaxError):
            continue
        rel = path.relative_to(ROOT).as_posix()
        lines = src.splitlines()
        for node in ast.walk(tree):
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)) and node.name.startswith(
                "test_"
            ):
                doc = ast.get_docstring(node)
                out[(rel, node.name)] = {
                    "section": _section_at_line(lines, node.lineno),
                    "description": (
                        doc.strip()
                        if doc
                        else "_(No docstring — add one under the def line in the test file.)_"
                    ),
                }
    return out


def _nodeid_function_name(nodeid: str) -> str:
    last = nodeid.split("::")[-1]
    return _base_test_name(last)


def _suite_for_path(relpath: str) -> str:
    if relpath.startswith("tests/parity/"):
        return "Parity"
    if relpath.startswith("tests/ui/"):
        return "UI"
    if relpath.startswith("tests/integration/"):
        return "Integration"
    if relpath.startswith("tests/unit/"):
        return "Unit"
    if relpath.startswith("tests/benchmarks/"):
        return "Benchmarks"
    return "Tests"


def _collect_nodeids(pytest_args: list[str] | None = None) -> tuple[list[str], str | None]:
    py, py_err = select_pytest_interpreter(ROOT)
    if py is None:
        return [], py_err or "No Python interpreter available for pytest."
    cmd = [py, "-m", "pytest", "--collect-only", "-q", *(pytest_args or [str(ROOT / "tests")])]
    # Reviewed: command is constructed from local interpreter + pytest args.
    proc = subprocess.run(  # nosec B603
        cmd,
        cwd=str(ROOT),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    nodeids: list[str] = []
    for raw in (proc.stdout or "").splitlines():
        line = raw.strip()
        if not line or COLLECT_TAIL_RE.match(line):
            continue
        if "::" not in line:
            continue
        nodeids.append(line)
    if proc.returncode != 0:
        tail = (proc.stdout or "")[-4000:] + "\n" + (proc.stderr or "")[-4000:]
        return nodeids, tail.strip() or f"pytest collect exited {proc.returncode}"
    return nodeids, None


def _test_collection_fingerprint() -> tuple[tuple[str, int, int], ...]:
    """Hashable file-state snapshot for invalidating cached pytest collection."""
    paths = [ROOT / "pytest.ini", ROOT / "pyproject.toml"]
    paths.extend(sorted((ROOT / "tests").rglob("test_*.py")))
    out: list[tuple[str, int, int]] = []
    for path in paths:
        try:
            stt = path.stat()
        except OSError:
            continue
        out.append((path.relative_to(ROOT).as_posix(), int(stt.st_mtime_ns), int(stt.st_size)))
    return tuple(out)


@st.cache_data(show_spinner=False)
def _cached_discover_tests_metadata_with_error(
    pytest_args: tuple[str, ...],
    fingerprint: tuple[tuple[str, int, int], ...],
) -> tuple[list[dict[str, Any]], str | None]:
    """Cached collection for Streamlit renders; invalidated by test file metadata."""
    del fingerprint
    return discover_tests_metadata_with_error(list(pytest_args))


def _discover_tests_metadata_for_render(
    pytest_args: list[str],
) -> tuple[list[dict[str, Any]], str | None]:
    if not hasattr(st, "cache_data"):
        return discover_tests_metadata_with_error(pytest_args)
    return _cached_discover_tests_metadata_with_error(
        tuple(pytest_args),
        _test_collection_fingerprint(),
    )


def _metadata_rows_from_nodeids(nodeids: list[str]) -> list[dict[str, Any]]:
    docs = _docstring_index()
    rows: list[dict[str, Any]] = []
    for nodeid in nodeids:
        relpath = nodeid.split("::", 1)[0]
        func = _nodeid_function_name(nodeid)
        meta = docs.get((relpath, func), {})
        suite = _suite_for_path(relpath)
        section = meta.get("section") or suite
        rows.append(
            {
                "nodeid": nodeid,
                "name": nodeid.split("::")[-1],
                "base_name": func,
                "file": relpath,
                "suite": suite,
                "section": section if section != "General" else suite,
                "description": meta.get(
                    "description",
                    "_(No docstring — add one under the def line in the test file.)_",
                ),
            }
        )
    return rows


def discover_tests_metadata_with_error(
    pytest_args: list[str] | None = None,
) -> tuple[list[dict[str, Any]], str | None]:
    """Discover pytest metadata and preserve collection diagnostics for the UI."""
    nodeids, err = _collect_nodeids(pytest_args)
    return _metadata_rows_from_nodeids(nodeids), err


def discover_tests_metadata(pytest_args: list[str] | None = None) -> list[dict[str, Any]]:
    """Discover pytest nodeids and enrich them with docstrings when available."""
    rows, err = discover_tests_metadata_with_error(pytest_args)
    if err and not rows:
        return []
    return rows


def run_pytest_junit(pytest_args: list[str] | None = None) -> tuple[int, str]:
    """Run pytest; write JUnit XML. Returns (exit_code, stderr+stdout snippet)."""
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    py, py_err = select_pytest_interpreter(ROOT)
    if py is None:
        msg = py_err or "No Python interpreter available for pytest."
        return 2, msg
    cmd = [
        py,
        "-m",
        "pytest",
        *(pytest_args or [str(ROOT / "tests")]),
        "-v",
        "--tb=short",
        f"--junitxml={JUNIT_PATH}",
        "-o",
        "junit_family=xunit2",
    ]
    # Reviewed: command is constructed from local interpreter + pytest args.
    proc = subprocess.run(  # nosec B603
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


def _render_pytest_unavailable(setup_error: str) -> None:
    st.warning("Unit tests require a pytest-capable interpreter.")
    st.markdown(
        "This environment is missing the minimal test dependencies needed to collect and run "
        "the in-app Unit Tests tab. On Streamlit Cloud, confirm the root `requirements.txt` "
        "was installed; locally, install the project runtime and dev dependency files."
    )
    st.markdown("Local setup:")
    st.code(LOCAL_PYTEST_SETUP, language="bash")
    with st.expander("Interpreter details"):
        st.code(setup_error, language="text")


def _render_collection_failure(collect_err: str | None, *, filter_text: str) -> None:
    if filter_text:
        st.warning("No tests matched the selected suite and filter.")
    else:
        st.error("Pytest collection did not return any tests.")
    if collect_err:
        with st.expander("pytest collection output"):
            st.code(collect_err, language="text")


def _nodeid_from_junit_case(tc: ElementTree.Element) -> str:
    file_attr = tc.get("file")
    name = tc.get("name") or ""
    if file_attr:
        rel = Path(file_attr).as_posix()
        return f"{rel}::{name}"
    classname = (tc.get("classname") or "").replace(".", "/")
    if classname:
        # pytest's xunit2 output usually omits ``file`` and stores a module-like
        # classname. Convert it back to the repository's nodeid shape.
        rel = f"{classname}.py"
        return f"{rel}::{name}"
    return name


def parse_junit_results() -> dict[str, dict[str, Any]]:
    """
    Map pytest nodeid -> {status, message, time_s}.
    status in passed | failed | skipped | error | unknown.

    Parametrized tests are preserved as individual nodeids so the dashboard's
    display matches pytest's collected test count exactly.
    """
    if not JUNIT_PATH.is_file():
        return {}
    try:
        tree = ElementTree.parse(JUNIT_PATH)
    except ElementTree.ParseError:
        return {}
    root = tree.getroot()

    raw: dict[str, list[dict[str, Any]]] = {}
    for tc in root.iter("testcase"):
        nodeid = _nodeid_from_junit_case(tc)
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
        raw.setdefault(nodeid, []).append(entry)

    out: dict[str, dict[str, Any]] = {}
    for nodeid, entries in raw.items():
        worst = max(entries, key=lambda e: _STATUS_SEVERITY.get(e["status"], 0))
        messages = [e["message"] for e in entries if e["message"]]
        total_time = None
        try:
            total_time = str(
                round(sum(float(e["time_s"]) for e in entries if e["time_s"] is not None), 4)
            )
        except (TypeError, ValueError):
            pass
        out[nodeid] = {
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
        st.title("Model unit test dashboard")
        st.caption(
            "Each row is one automated check of `pricing_projection.py`. "
            "Discovery is powered by `pytest --collect-only`, so it reflects the full project suite."
        )
    else:
        st.subheader("Unit tests")
        st.caption(
            "Automated checks collected from the full pytest suite. Use the filters to run a smaller gate."
        )

    selected_preset = st.session_state.get("pytest_preset", "Full suite")
    filter_text = st.session_state.get("pytest_filter", "")
    run_args = list(PRESET_ARGS.get(selected_preset, PRESET_ARGS["Full suite"]))
    if filter_text:
        run_args.extend(["-k", filter_text])

    _py_ok, py_setup_err = select_pytest_interpreter(ROOT)
    if py_setup_err:
        _render_pytest_unavailable(py_setup_err)
        return

    meta, collect_err = _discover_tests_metadata_for_render(run_args)
    if not meta:
        _render_collection_failure(collect_err, filter_text=filter_text)
        return
    if collect_err:
        st.warning("pytest reported collection diagnostics; showing the tests it did collect.")
        with st.expander("pytest collection output"):
            st.code(collect_err, language="text")

    notify = st.session_state.get("last_notify")
    if notify == "pass":
        st.success("Last test run finished with pytest exit code 0 (all executed tests passed).")
    elif notify == "fail":
        st.warning(
            "Last test run reported failures, errors, or a non-zero pytest exit code. See expanders below."
        )

    def _run_clicked() -> None:
        with st.spinner("Running pytest…"):
            code, log_tail = run_pytest_junit(run_args)
        st.session_state["last_exit_code"] = code
        st.session_state["last_log_tail"] = log_tail
        st.session_state["last_results"] = parse_junit_results()
        st.session_state["last_notify"] = "pass" if code == 0 else "fail"

    if embedded:
        c_preset, c_filter, c_run = st.columns([1.2, 1.5, 1.0])
        with c_preset:
            st.selectbox(
                "Suite",
                list(PRESET_ARGS),
                key="pytest_preset",
                label_visibility="collapsed",
            )
        with c_filter:
            st.text_input(
                "Filter",
                key="pytest_filter",
                placeholder="-k expression",
                label_visibility="collapsed",
            )
        with c_run:
            if st.button(
                "Run selected", type="primary", width="stretch", key="pytest_run_embedded"
            ):
                _run_clicked()
    else:
        with st.sidebar:
            st.header("Run")
            st.selectbox("Suite", list(PRESET_ARGS), key="pytest_preset")
            st.text_input("Filter", key="pytest_filter", placeholder="-k expression")
            if st.button("Run selected", type="primary", width="stretch"):
                _run_clicked()
            st.divider()
            st.markdown(
                "**First time setup:** in a terminal here, run  \n`python -m pip install -r requirements-dev.txt`"
            )
            st.markdown(
                "**CLI alternative:** `python -m pytest` or `run_tests_report.bat` for HTML."
            )

    results: dict[str, dict[str, Any]] = st.session_state.get("last_results") or {}

    # Summary metrics
    nodeids = [m["nodeid"] for m in meta]
    passed = sum(1 for n in nodeids if results.get(n, {}).get("status") == "passed")
    failed = sum(1 for n in nodeids if results.get(n, {}).get("status") in ("failed", "error"))
    skipped = sum(1 for n in nodeids if results.get(n, {}).get("status") == "skipped")
    not_run = sum(1 for n in nodeids if n not in results)

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total tests", len(meta))
    c2.metric("Passed", passed)
    c3.metric("Failed", failed)
    c4.metric("Skipped", skipped)
    c5.metric("Not run yet", not_run)

    st.divider()

    # Group by section
    st.caption(
        f"Showing `{selected_preset}` tests"
        + (f" matching `-k {filter_text}`." if filter_text else ".")
    )

    sections: dict[str, list[dict[str, Any]]] = {}
    for m in meta:
        sections.setdefault(f"{m['suite']} / {m['section']}", []).append(m)

    for section in sorted(sections.keys()):
        st.subheader(section)
        for m in sections[section]:
            nodeid = m["nodeid"]
            r = results.get(nodeid, {})
            status = r.get("status", "not_run")
            icon = {
                "passed": "[pass]",
                "failed": "[fail]",
                "error": "[error]",
                "skipped": "[skip]",
                "not_run": "[ ]",
            }.get(status, "[ ]")
            with st.expander(
                f"{icon} **{m['name']}** — _{status.replace('_', ' ')}_",
                expanded=(status in ("failed", "error")),
            ):
                st.code(nodeid, language="text")
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
    st.set_page_config(page_title="Model unit tests", layout="wide")
    render_unit_tests_page(embedded=False)


if __name__ == "__main__":
    main()
