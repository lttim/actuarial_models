"""Shared helpers for Streamlit AppTest smoke tests.

These tests exercise ``pricing_ui.py`` through Streamlit's official
``streamlit.testing.v1.AppTest`` harness. The point is *not* to assert
numerical behaviour (that is what ``tests/parity/`` and
``tests/test_regression_matrix.py`` are for); the point is to catch the
class of bug that only surfaces when the Streamlit script is actually
executed end-to-end:

* a missing ``st.session_state`` initialisation that raises ``KeyError``
  on first paint,
* a widget key that no longer matches a constant in
  ``pricing_run_form_state``,
* an Excel-download path that constructs an invalid workbook,
* a product selectbox option that is no longer routable to a valid
  pricing engine.

Each per-product test follows the same shape:

  1. Load the app via :func:`load_pricing_ui`.
  2. Navigate to the ``Pricing Run`` page (``run`` section).
  3. Set the product selectbox to the product under test.
  4. Re-run.
  5. Assert no exception was raised AND the product-specific contract
     widgets are visible.

Helpers below absorb the boilerplate so each per-product test is a
single, readable function.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


@pytest.fixture(scope="module")
def streamlit_apptest_module():
    """Import ``streamlit.testing.v1`` lazily so collection never fails.

    AppTest is sensitive to the streamlit version (introduced in 1.28);
    if it is not available, every UI test in the module is skipped with
    a clear reason rather than raising at collection time.
    """
    try:
        from streamlit.testing.v1 import AppTest
    except ImportError as exc:
        pytest.skip(f"streamlit.testing.v1 not available: {exc!r}")
    return AppTest


def load_pricing_ui(AppTest, default_timeout: int = 120):
    """Load ``pricing_ui.py`` and return the AppTest instance after one run.

    Raises immediately if the initial render raised an exception; that
    is a hard failure of the smoke contract regardless of which product
    the test ultimately exercises.

    The default timeout is intentionally above Streamlit's 60s default because
    the full coverage gate instruments workbook-building paths and can push a
    legitimate pricing-click rerun just past 60s on local machines.
    """
    app_path = ROOT / "pricing_ui.py"
    at = AppTest.from_file(str(app_path), default_timeout=default_timeout)
    at.run()
    if at.exception:
        # Surface the real exception text so the failure is actionable
        # rather than the generic "exception count = N".
        raise AssertionError(
            "pricing_ui.py raised on initial render:\n"
            + "\n---\n".join(str(e.value) for e in at.exception)
        )
    return at


def navigate_to_pricing_run(at) -> None:
    """Click the sidebar 'Pricing Run' radio option and re-run the app.

    Streamlit testing addresses radios positionally inside a section. The
    sidebar radio in ``pricing_ui.main()`` is the ONLY radio rendered
    before any user interaction (the per-page radios live deeper in the
    DOM and don't appear until the corresponding page is opened), so
    ``at.radio[0]`` is the section radio.
    """
    assert at.radio, "expected at least one radio (sidebar Section radio)"
    section_radio = at.radio[0]
    # The internal storage uses the option keys, not the labels.
    # Try-set by value; SECTION_ORDER includes 'run'.
    section_radio.set_value("run").run()


def select_product(at, product_value: str) -> None:
    """Set the ``run_product_type`` selectbox and re-run."""
    matched = [s for s in at.selectbox if s.key == "run_product_type"]
    assert matched, (
        "Pricing Run page is missing the 'run_product_type' selectbox; "
        "either the page failed to render or the key was renamed in "
        "pricing_ui.py without updating the AppTest smoke tests."
    )
    matched[0].set_value(product_value).run()


def assert_no_exceptions(at, *, context: str) -> None:
    """Reusable terminal assertion: no script-level exceptions."""
    if not at.exception:
        return
    raise AssertionError(
        f"AppTest exceptions during {context}:\n"
        + "\n---\n".join(str(e.value) for e in at.exception)
    )
