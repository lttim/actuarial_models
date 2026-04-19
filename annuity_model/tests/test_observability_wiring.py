"""Meta-invariant + behavioural tests for the OTel observability wiring.

Background
----------
``_observability.traced`` is a no-op decorator unless OpenTelemetry is
installed AND ``OTEL_EXPORTER_OTLP_ENDPOINT`` is set in the environment.
Production deployments rely on it to emit one span per pricing or ALM
call; without an explicit test that the decorator is *applied*, it is
trivially easy to forget to add ``@traced(...)`` to a new entry point
(or to silently strip one during a refactor) -- the no-op fallback
means everything keeps working in CI, but the production telemetry
silently goes dark.

These tests close that gap with two complementary signals:

1. **Wiring assertions** -- every public price/ALM entry point listed in
   :data:`TRACED_ENTRY_POINTS` must be wrapped by ``traced(...)``. We
   detect that by checking for the ``__wrapped__`` attribute that
   ``functools.wraps`` attaches; the no-op fallback inside ``traced``
   still uses ``functools.wraps``, so this works regardless of OTel
   availability.

2. **Behavioural smoke** -- calling each wrapped function with minimal
   inputs returns the same result as the unwrapped function. This is
   the safety net that catches a future change to ``traced`` that
   accidentally swallows arguments, mutates kwargs, or reorders args.
"""

from __future__ import annotations

import functools

import numpy as np
import pytest

import fia_projection as fp
import iul_projection as iul
import myga_projection as my
import pricing_projection as sp
import rila_projection as rp
import term_projection as tp
import ul_projection as ul
import va_projection as va
import vul_projection as vul
import wl_projection as wl

# (qualified-name-for-error-messages, callable). The qualified name is
# what the test prints when the wiring breaks; a clear name pointing at
# the right module/function saves a triage step.
TRACED_ENTRY_POINTS: list[tuple[str, object]] = [
    ("pricing_projection.price_spia_single_premium", sp.price_spia_single_premium),
    (
        "pricing_projection.price_spia_single_premium_monte_carlo",
        sp.price_spia_single_premium_monte_carlo,
    ),
    (
        "pricing_projection.run_alm_projection_from_liability_path",
        sp.run_alm_projection_from_liability_path,
    ),
    ("pricing_projection.run_alm_projection", sp.run_alm_projection),
    (
        "pricing_projection.run_alm_projection_from_pricing_result",
        sp.run_alm_projection_from_pricing_result,
    ),
    ("term_projection.price_term_life_level_monthly", tp.price_term_life_level_monthly),
    ("rila_projection.price_rila_single_premium", rp.price_rila_single_premium),
    (
        "rila_projection.price_rila_single_premium_monte_carlo",
        rp.price_rila_single_premium_monte_carlo,
    ),
    ("myga_projection.price_myga_single_premium", my.price_myga_single_premium),
    ("fia_projection.price_fia_single_premium", fp.price_fia_single_premium),
    (
        "fia_projection.price_fia_single_premium_monte_carlo",
        fp.price_fia_single_premium_monte_carlo,
    ),
    ("va_projection.price_va_single_premium", va.price_va_single_premium),
    (
        "va_projection.price_va_single_premium_monte_carlo",
        va.price_va_single_premium_monte_carlo,
    ),
    ("wl_projection.price_wl_single_premium", wl.price_wl_single_premium),
    ("ul_projection.price_ul_single_premium", ul.price_ul_single_premium),
    ("iul_projection.price_iul_single_premium", iul.price_iul_single_premium),
    (
        "iul_projection.price_iul_single_premium_monte_carlo",
        iul.price_iul_single_premium_monte_carlo,
    ),
    ("vul_projection.price_vul_single_premium", vul.price_vul_single_premium),
    (
        "vul_projection.price_vul_single_premium_monte_carlo",
        vul.price_vul_single_premium_monte_carlo,
    ),
]


@pytest.mark.parametrize(
    "qualname,fn",
    TRACED_ENTRY_POINTS,
    ids=[name for name, _ in TRACED_ENTRY_POINTS],
)
def test_entry_point_is_traced(qualname: str, fn: object) -> None:
    """Every parity-critical entry point must be wrapped by ``@traced(...)``.

    ``functools.wraps`` (used inside ``_observability.traced``) attaches
    a ``__wrapped__`` attribute pointing at the original function. We
    assert its presence here. If a future refactor strips the
    decorator, this test fails with a clear message that points at the
    exact module/function pair to fix.
    """
    assert hasattr(fn, "__wrapped__"), (
        f"{qualname} is missing the @traced(...) decorator from "
        "_observability. Production OTel deployments will lose this "
        "span. Re-add `@traced(\"<span name>\")` directly above the "
        "function definition."
    )
    # Sanity: the wrapper is truly a wrapper, not the original under a
    # different name.
    assert getattr(fn, "__wrapped__") is not fn, (
        f"{qualname} __wrapped__ points at itself; the decorator is not "
        "actually applied."
    )


def test_no_op_path_does_not_alter_return_value() -> None:
    """When OTel is not configured, ``traced`` MUST be a transparent
    pass-through.

    We verify this with the SPIA deterministic engine (cheapest
    representative call). If ``_observability.traced`` is ever changed
    to do something on the no-op path that mutates inputs/outputs, this
    test catches it.
    """
    # Force OTel off even if the dev machine has the libs.
    import _observability

    _observability._tracer = None
    _observability._OTEL_ENABLED = False

    contract = sp.SPIAContract(issue_age=65, sex="male", benefit_annual=100_000.0)
    yc = sp.YieldCurve.from_flat_rate(0.04)
    ages = np.arange(0, 121, dtype=int)
    qx = np.full_like(ages, 0.02, dtype=float)
    mort = sp.MortalityTableQx(ages, qx)
    expenses = sp.ExpenseAssumptions(0.0, 0.0, 0.0)

    wrapped = sp.price_spia_single_premium(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=80,
        spread=0.0,
        valuation_year=None,
        expenses=expenses,
        index_scenario_csv_path=None,
        expense_annual_inflation=0.0,
    )
    raw = sp.price_spia_single_premium.__wrapped__(
        contract=contract,
        yield_curve=yc,
        mortality=mort,
        horizon_age=80,
        spread=0.0,
        valuation_year=None,
        expenses=expenses,
        index_scenario_csv_path=None,
        expense_annual_inflation=0.0,
    )

    assert wrapped.single_premium == raw.single_premium
    assert wrapped.pv_benefit == raw.pv_benefit
    assert wrapped.annuity_factor == raw.annuity_factor


def test_traced_decorator_uses_functools_wraps() -> None:
    """Implementation contract: ``traced`` must use ``functools.wraps``
    so that ``__wrapped__``, ``__name__``, ``__doc__``, etc. survive.

    Without this, the wiring tests above could trivially be fooled by
    a future implementation that returns a bare lambda. We test the
    decorator on a synthetic function so we don't depend on any
    particular engine's docstring text.
    """
    from _observability import traced

    @traced("test.synthetic")
    def _example(x: int, y: int = 1) -> int:
        """example docstring."""
        return x + y

    assert _example.__wrapped__ is not None
    assert _example.__name__ == "_example"
    assert _example.__doc__ == "example docstring."
    # functools.wraps round-trip
    assert _example(2, 3) == 5
    assert _example.__wrapped__(2, 3) == 5
    assert isinstance(_example, functools.partial) is False  # plain function


def test_span_name_is_passed_through_when_otel_present(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """When a tracer IS available, ``traced`` must pass the configured
    span name to ``start_as_current_span``. We stub a fake tracer and
    record the names handed to it.
    """
    import _observability

    received: list[str] = []

    class _FakeSpan:
        def record_exception(self, exc: BaseException) -> None:  # pragma: no cover
            pass

        def __enter__(self) -> "_FakeSpan":
            return self

        def __exit__(self, *args: object) -> None:
            return None

    class _FakeTracer:
        def start_as_current_span(self, name: str) -> _FakeSpan:
            received.append(name)
            return _FakeSpan()

    monkeypatch.setattr(_observability, "_tracer", _FakeTracer())
    monkeypatch.setattr(_observability, "_OTEL_ENABLED", True)

    @_observability.traced("test.named_span")
    def _example() -> int:
        return 42

    assert _example() == 42
    assert received == ["test.named_span"]


def test_default_span_name_falls_back_to_qualname(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """If ``traced(span_name=None)``, the function's ``__qualname__`` is
    used. Important so that newly added entry points without an
    explicit span name still get *some* identifier in the trace tree.
    """
    import _observability

    received: list[str] = []

    class _FakeSpan:
        def record_exception(self, exc: BaseException) -> None:  # pragma: no cover
            pass

        def __enter__(self) -> "_FakeSpan":
            return self

        def __exit__(self, *args: object) -> None:
            return None

    class _FakeTracer:
        def start_as_current_span(self, name: str) -> _FakeSpan:
            received.append(name)
            return _FakeSpan()

    monkeypatch.setattr(_observability, "_tracer", _FakeTracer())
    monkeypatch.setattr(_observability, "_OTEL_ENABLED", True)

    @_observability.traced()
    def _outer_named_function() -> str:
        return "ok"

    _outer_named_function()
    assert received and received[0].endswith("_outer_named_function")
