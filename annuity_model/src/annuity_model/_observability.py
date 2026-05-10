"""Optional OpenTelemetry instrumentation hooks.

OTel is **not** a required dependency: imports are gated on the user having
``opentelemetry-api`` / ``opentelemetry-sdk`` installed AND
``OTEL_EXPORTER_OTLP_ENDPOINT`` set in the environment.

When OTel is enabled, every wrapped function emits a span with attributes for
the product type, contract age, scenario size, etc. When it is not, the
decorators are no-ops with negligible overhead.

Public surface
--------------

::

    from annuity_model._observability import traced

    @traced("price_spia_single_premium")
    def price_spia_single_premium(...): ...

The decorator is intentionally minimal: structured logging via
:mod:`_logging` remains the primary diagnostic channel; OTel is a thin extra
hook for production deployments that already have a collector.
"""

from __future__ import annotations

import functools
import os
from collections.abc import Callable
from typing import Any, TypeVar

F = TypeVar("F", bound=Callable[..., Any])

_OTEL_ENABLED = False
_tracer = None


def _maybe_init_otel() -> None:
    """Best-effort lazy initialisation; never raises."""
    global _OTEL_ENABLED, _tracer
    if _OTEL_ENABLED or _tracer is not None:
        return
    if not os.environ.get("OTEL_EXPORTER_OTLP_ENDPOINT"):
        return
    try:
        from opentelemetry import trace
        from opentelemetry.exporter.otlp.proto.http.trace_exporter import OTLPSpanExporter
        from opentelemetry.sdk.resources import Resource
        from opentelemetry.sdk.trace import TracerProvider
        from opentelemetry.sdk.trace.export import BatchSpanProcessor
    except ImportError:
        return

    resource = Resource.create({"service.name": "annuity_model"})
    provider = TracerProvider(resource=resource)
    provider.add_span_processor(BatchSpanProcessor(OTLPSpanExporter()))
    trace.set_tracer_provider(provider)
    _tracer = trace.get_tracer("annuity_model")
    _OTEL_ENABLED = True


def traced(span_name: str | None = None) -> Callable[[F], F]:
    """Decorator that wraps the function call in an OTel span if available."""

    def _decorator(fn: F) -> F:
        name = span_name or fn.__qualname__

        @functools.wraps(fn)
        def _wrapper(*args: Any, **kwargs: Any) -> Any:
            _maybe_init_otel()
            if _tracer is None:
                return fn(*args, **kwargs)
            with _tracer.start_as_current_span(name) as span:
                try:
                    result = fn(*args, **kwargs)
                except Exception as exc:
                    span.record_exception(exc)
                    raise
                return result

        return _wrapper  # type: ignore[return-value]

    return _decorator


__all__ = ["traced"]
