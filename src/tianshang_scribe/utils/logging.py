"""Structured logging configuration built on ``structlog``.

Replaces ad-hoc ``print``/standard-library logging in the MCP server with
context-rich, machine-parseable events. In default (console) mode output is
human-friendly coloured text; setting ``TIANSHANG_SCRIBE_LOG_JSON=1`` (or
``log_json=True``) switches to JSON lines for ELK/Loki-style pipelines. Every
event carries a timestamp, level, and event name plus caller-supplied key
values.
"""

from __future__ import annotations

import logging
from typing import Any

import structlog

from tianshang_scribe.utils.config import Settings


def get_logger(name: str = 'scribe') -> structlog.stdlib.BoundLogger:
    """Return a bound logger configured with the process-wide settings."""
    logger: structlog.stdlib.BoundLogger = structlog.get_logger(name)
    return logger


def configure_logging(settings: Settings | None = None) -> None:
    """Configure structlog processors, level and output format globally.

    ``settings`` may be supplied explicitly (e.g. from a CLI-parsed config);
    otherwise a fresh :class:`Settings` is read so ``TIANSHANG_SCRIBE_LOG_*`` env vars
    and a ``.env`` file apply. Safe to call more than once.
    """
    settings = settings or Settings()

    level = getattr(logging, settings.log_level.upper(), logging.INFO)
    renderer: structlog.types.Processor = (
        structlog.processors.JSONRenderer()
        if settings.log_json
        else structlog.dev.ConsoleRenderer()
    )

    structlog.configure(
        processors=[
            structlog.contextvars.merge_contextvars,
            structlog.processors.add_log_level,
            structlog.processors.TimeStamper(fmt='iso'),
            renderer,
        ],
        wrapper_class=structlog.make_filtering_bound_logger(level),
        cache_logger_on_first_use=True,
    )

    # Route standard-library records (e.g. uvicorn access logs) through the
    # structlog pipeline so server output stays homogeneous.
    logging.basicConfig(level=level, force=True)


def log_http_event(
    event: str,
    *,
    method: str,
    path: str,
    status_code: int | None = None,
    **extra: Any,
) -> None:
    """Emit a structured HTTP event with common request context."""
    get_logger('scribe.http').info(
        event,
        method=method,
        path=path,
        status_code=status_code,
        **extra,
    )
