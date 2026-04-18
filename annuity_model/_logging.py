"""Structured logging for annuity_model.

Replaces the historical ``print()``-based diagnostics with a real
:mod:`logging` setup that:

* Emits human-readable lines to stderr in dev mode (default).
* Emits one JSON object per line when ``ANNUITY_MODEL_LOG_FORMAT=json`` so
  CI / Docker can pipe the stream into log aggregators (Datadog, Loki, ...).
* Honours ``ANNUITY_MODEL_LOG_LEVEL`` (default ``INFO``) and a per-logger
  override env ``ANNUITY_MODEL_LOG_<MODULE>=DEBUG``.

Public API
----------

::

    from annuity_model import get_logger, configure_logging
    configure_logging()                # idempotent; safe to call from any entry point
    log = get_logger(__name__)
    log.info("priced %s", contract)
    log.warning("validator surfaced %d issues", len(issues))

Existing ``print()`` sites in :mod:`pricing_projection`,
:mod:`build_pricing_excel_workbook`, and the illustration scripts are being
migrated incrementally; new code MUST use the logger so that future log
routing changes (file sinks, OTel exporters in Phase 4) flow through one
configuration.
"""

from __future__ import annotations

import json
import logging
import os
import sys
import time
from typing import Any

_LOGGER_PREFIX = "annuity_model"
_DEFAULT_LEVEL = "INFO"
_ENV_LEVEL = "ANNUITY_MODEL_LOG_LEVEL"
_ENV_FORMAT = "ANNUITY_MODEL_LOG_FORMAT"  # "text" (default) or "json"
_ENV_PER_LOGGER_PREFIX = "ANNUITY_MODEL_LOG_"  # e.g. ANNUITY_MODEL_LOG_PRICING_PROJECTION=DEBUG

_configured = False


class _JsonLineFormatter(logging.Formatter):
    """Render every record as a single-line JSON object.

    Preserves ``extra=`` keyword args for structured fields without forcing
    callers to format them into the message string.
    """

    _RESERVED = frozenset({
        "name", "msg", "args", "levelname", "levelno", "pathname", "filename",
        "module", "exc_info", "exc_text", "stack_info", "lineno", "funcName",
        "created", "msecs", "relativeCreated", "thread", "threadName",
        "processName", "process", "message", "asctime", "taskName",
    })

    def format(self, record: logging.LogRecord) -> str:
        payload: dict[str, Any] = {
            "ts": time.strftime("%Y-%m-%dT%H:%M:%S", time.gmtime(record.created))
                  + f".{int(record.msecs):03d}Z",
            "level": record.levelname,
            "logger": record.name,
            "msg": record.getMessage(),
        }
        if record.exc_info:
            payload["exc"] = self.formatException(record.exc_info)
        # Surface any extra= fields the caller passed in.
        for k, v in record.__dict__.items():
            if k in self._RESERVED or k.startswith("_"):
                continue
            try:
                json.dumps(v)
                payload[k] = v
            except (TypeError, ValueError):
                payload[k] = repr(v)
        return json.dumps(payload, default=str, ensure_ascii=False)


def _resolve_level(name: str | None) -> int:
    raw = (name or "").strip().upper()
    if not raw:
        return logging.INFO
    if raw.isdigit():
        return int(raw)
    return logging.getLevelNamesMapping().get(raw, logging.INFO)


def configure_logging(
    *,
    level: str | int | None = None,
    fmt: str | None = None,
    stream: Any = None,
    force: bool = False,
) -> None:
    """Idempotently configure the ``annuity_model`` logger family.

    Call once near the top of any entry-point script (Streamlit, deep_smoke,
    illustration drivers). A second call without ``force=True`` is a no-op so
    libraries can call it defensively without trampling caller config.

    Parameters
    ----------
    level:
        Override the level for the root ``annuity_model`` logger. Falls back
        to ``$ANNUITY_MODEL_LOG_LEVEL`` then ``INFO``.
    fmt:
        ``"text"`` (default) or ``"json"``. Falls back to
        ``$ANNUITY_MODEL_LOG_FORMAT`` then ``"text"``.
    stream:
        Override the destination stream (default ``sys.stderr``).
    force:
        Reconfigure even if previously initialized.
    """
    global _configured
    if _configured and not force:
        return

    base = logging.getLogger(_LOGGER_PREFIX)
    base.handlers.clear()
    base.propagate = False

    resolved_level = _resolve_level(
        str(level) if level is not None else os.environ.get(_ENV_LEVEL, _DEFAULT_LEVEL)
    )
    base.setLevel(resolved_level)

    handler = logging.StreamHandler(stream or sys.stderr)
    handler.setLevel(resolved_level)

    chosen_fmt = (fmt or os.environ.get(_ENV_FORMAT, "text")).strip().lower()
    if chosen_fmt == "json":
        handler.setFormatter(_JsonLineFormatter())
    else:
        handler.setFormatter(logging.Formatter(
            fmt="%(asctime)s %(levelname)-7s %(name)s :: %(message)s",
            datefmt="%H:%M:%S",
        ))
    base.addHandler(handler)

    # Per-logger overrides via env: ANNUITY_MODEL_LOG_PRICING_PROJECTION=DEBUG
    for env_key, env_val in os.environ.items():
        if not env_key.startswith(_ENV_PER_LOGGER_PREFIX):
            continue
        if env_key in (_ENV_LEVEL, _ENV_FORMAT):
            continue
        suffix = env_key[len(_ENV_PER_LOGGER_PREFIX):].lower()
        if not suffix:
            continue
        logging.getLogger(f"{_LOGGER_PREFIX}.{suffix}").setLevel(_resolve_level(env_val))

    _configured = True


def get_logger(name: str) -> logging.Logger:
    """Return a logger under the ``annuity_model.*`` namespace.

    ``name`` is normalised: an absolute module path like
    ``"annuity_model.engines.spia"`` is taken as-is, while ``"__main__"``,
    ``"pricing_projection"``, etc. are placed under the ``annuity_model``
    prefix so they share configuration.
    """
    if not name or name == "__main__":
        full = _LOGGER_PREFIX
    elif name == _LOGGER_PREFIX or name.startswith(f"{_LOGGER_PREFIX}."):
        full = name
    else:
        full = f"{_LOGGER_PREFIX}.{name}"
    return logging.getLogger(full)


__all__ = ["configure_logging", "get_logger"]
