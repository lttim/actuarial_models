"""Streamlit Cloud runtime smoke.

This intentionally emulates Streamlit Community Cloud's production install
surface: root ``requirements.txt`` plus root ``streamlit_app.py``. It must not
depend on an editable package install, ``requirements.lock``, or dev tools.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

REPO_ROOT = Path(__file__).resolve().parents[2]
SRC_ROOT = REPO_ROOT / "annuity_model" / "src"
CLOUD_ENTRY = REPO_ROOT / "streamlit_app.py"
STARTUP_ERROR = "The app failed while starting."


def _element_values(elements: Any) -> list[str]:
    values: list[str] = []
    for element in elements or []:
        value = getattr(element, "value", None)
        if value is None:
            value = str(element)
        values.append(str(value))
    return values


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

    exception_values = _element_values(getattr(at, "exception", []))
    error_values = _element_values(getattr(at, "error", []))
    all_error_text = "\n".join(exception_values + error_values)
    if exception_values or STARTUP_ERROR in all_error_text:
        print("FAIL: streamlit_app.py did not boot cleanly.", file=sys.stderr)
        if exception_values:
            print("AppTest exceptions:", file=sys.stderr)
            print("\n---\n".join(exception_values), file=sys.stderr)
        if error_values:
            print("Streamlit errors:", file=sys.stderr)
            print("\n---\n".join(error_values), file=sys.stderr)
        return 1

    print("PASS: streamlit_app.py boots under the Streamlit Cloud runtime surface.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
