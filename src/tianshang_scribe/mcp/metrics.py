"""Prometheus metrics for the TianshangScribe MCP Server.

Exposes operation duration/count histograms in the Prometheus text format.
Tool dispatch is instrumented via :func:`track_operation`; the metric labels
are bounded (``tool_name`` below 50, ``status`` in {success, error}).
"""

from __future__ import annotations

import functools
import time
from collections.abc import Callable, Iterator
from contextlib import contextmanager
from typing import Any, TypeVar

from prometheus_client import CONTENT_TYPE_LATEST, Counter, Histogram, generate_latest
from starlette.responses import Response

F = TypeVar('F', bound=Callable[..., Any])

OPERATION_DURATION = Histogram(
    'tianshang_scribe_operation_duration_seconds',
    'Duration of TianshangScribe MCP tool operations',
    ['tool_name'],
    buckets=[0.001, 0.005, 0.01, 0.025, 0.05, 0.1, 0.25, 0.5, 1.0, 2.5, 5.0, 10.0],
)

OPERATION_COUNT = Counter(
    'tianshang_scribe_operations_total',
    'Total number of TianshangScribe MCP tool operations',
    ['tool_name', 'status'],
)


def observe_operation(tool_name: str) -> None:
    """Record a successful tool operation."""
    OPERATION_COUNT.labels(tool_name=tool_name, status='success').inc()


def observe_error(tool_name: str) -> None:
    """Record a failed tool operation."""
    OPERATION_COUNT.labels(tool_name=tool_name, status='error').inc()


@contextmanager
def track_operation(tool_name: str) -> Iterator[None]:
    """Time and record a tool operation, tagging success or error."""
    started = time.perf_counter()
    try:
        yield
        observe_operation(tool_name)
    except BaseException:
        observe_error(tool_name)
        raise
    finally:
        OPERATION_DURATION.labels(tool_name=tool_name).observe(time.perf_counter() - started)


def instrumented(fn: F, tool_name: str) -> F:
    """Wrap a tool function with :func:`track_operation`.

    ``functools.wraps`` preserves the function signature so the MCP SDK can
    still derive the ``inputSchema`` from it.
    """

    @functools.wraps(fn)
    def wrapper(*args: Any, **kwargs: Any) -> Any:
        with track_operation(tool_name):
            return fn(*args, **kwargs)

    return wrapper  # type: ignore[return-value]


def metrics_endpoint(request: object) -> Response:
    """Starlette route handler serving the Prometheus text format."""
    return Response(
        content=generate_latest(),
        media_type=CONTENT_TYPE_LATEST,
    )
