"""Streamlit Cloud runtime smoke.

This intentionally emulates Streamlit Community Cloud's production install
surface: root ``requirements.txt`` plus root ``streamlit_app.py``. It must not
depend on an editable package install, ``requirements.lock``, or the full dev
toolchain.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

REPO_ROOT = Path(__file__).resolve().parents[2]
SRC_ROOT = REPO_ROOT / "annuity_model" / "src"
CLOUD_ENTRY = REPO_ROOT / "streamlit_app.py"
STARTUP_ERROR = "The app failed while starting."
OLD_TEST_COLLECTION_ERROR = (
    "Could not collect tests. Open the `annuity_model` folder as project root and "
    "confirm pytest is installed in the selected interpreter."
)
PYTEST_DEV_ONLY_MESSAGE = "Unit tests are local development tooling"
PYTEST_UNAVAILABLE_MESSAGE = "Unit tests require a pytest-capable interpreter."


def _element_values(elements: Any) -> list[str]:
    values: list[str] = []
    for element in elements or []:
        value = getattr(element, "value", None)
        if value is None:
            value = str(element)
        values.append(str(value))
    return values


def _find_section_radio(at: Any) -> Any:
    for radio in at.radio:
        if getattr(radio, "label", None) == "Section":
            return radio
    labels = [getattr(radio, "label", None) for radio in at.radio]
    raise AssertionError(f"Could not find sidebar Section radio; saw labels {labels!r}.")


def _rendered_text(at: Any) -> str:
    groups = [
        getattr(at, "caption", []),
        getattr(at, "error", []),
        getattr(at, "exception", []),
        getattr(at, "markdown", []),
        getattr(at, "subheader", []),
        getattr(at, "warning", []),
    ]
    return "\n".join(value for group in groups for value in _element_values(group))


def _metric_value(at: Any, label: str) -> int | None:
    for metric in getattr(at, "metric", []):
        if getattr(metric, "label", None) == label:
            raw = getattr(metric, "value", None) or getattr(metric, "body", None)
            try:
                return int(str(raw).replace(",", ""))
            except (TypeError, ValueError):
                return None
    return None


def _assert_no_exception(at: Any, *, context: str) -> bool:
    exception_values = _element_values(getattr(at, "exception", []))
    error_values = _element_values(getattr(at, "error", []))
    all_error_text = "\n".join(exception_values + error_values)
    if exception_values or STARTUP_ERROR in all_error_text:
        print(f"FAIL: streamlit_app.py failed during {context}.", file=sys.stderr)
        if exception_values:
            print("AppTest exceptions:", file=sys.stderr)
            print("\n---\n".join(exception_values), file=sys.stderr)
        if error_values:
            print("Streamlit errors:", file=sys.stderr)
            print("\n---\n".join(error_values), file=sys.stderr)
        return False
    return True


def main() -> int:
    if str(SRC_ROOT) not in sys.path:
        sys.path.insert(0, str(SRC_ROOT))

    try:
        import annuity_model.pricing_ui  # noqa: F401
    except Exception as exc:
        print("FAIL: importing annuity_model.pricing_ui raised:", repr(exc), file=sys.stderr)
        return 1

    try:
        from streamlit.testing.v1 import AppTest
    except Exception as exc:
        print("FAIL: streamlit.testing.v1 is unavailable:", repr(exc), file=sys.stderr)
        return 1

    at = AppTest.from_file(str(CLOUD_ENTRY), default_timeout=120)
    at.run()

    if not _assert_no_exception(at, context="initial boot"):
        return 1

    try:
        _find_section_radio(at).set_value("tests").run()
    except Exception as exc:
        print("FAIL: navigating to Unit Tests raised:", repr(exc), file=sys.stderr)
        return 1

    if not _assert_no_exception(at, context="Unit Tests tab render"):
        return 1

    rendered_text = _rendered_text(at)
    if OLD_TEST_COLLECTION_ERROR in rendered_text:
        print("FAIL: Unit Tests tab showed the legacy pytest collection error.", file=sys.stderr)
        return 1
    if PYTEST_DEV_ONLY_MESSAGE in rendered_text or PYTEST_UNAVAILABLE_MESSAGE in rendered_text:
        print(
            "FAIL: Unit Tests tab did not collect tests under the Streamlit Cloud manifest.",
            file=sys.stderr,
        )
        return 1
    total_tests = _metric_value(at, "Total tests")
    if total_tests is None or total_tests <= 0:
        print(
            f"FAIL: Unit Tests tab did not render a positive collected test count: {total_tests!r}.",
            file=sys.stderr,
        )
        return 1

    print(
        f"PASS: streamlit_app.py boots and the Unit Tests tab collects {total_tests} tests under the Streamlit Cloud runtime surface."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
