"""HTTP transport wiring for the TianshangScribe MCP Server.

Builds on the official MCP SDK apps (:meth:`mcp.server.mcpserver.MCPServer.sse_app`
and ``.streamable_http_app``) and wraps them with CORS, bearer-token auth and
sliding-window rate limiting. ``/health`` and ``/metrics`` are registered on the
server itself via ``custom_route`` so they stay available on every transport.

Middleware is implemented as pure ASGI classes (not ``BaseHTTPMiddleware``) so
the long-lived SSE/streamable-HTTP GET streams are never buffered.
"""

from __future__ import annotations

import asyncio
import hmac
from collections.abc import Callable
from typing import Any

from starlette.middleware.cors import CORSMiddleware
from starlette.responses import JSONResponse, Response

from src.mcp.metrics import metrics_endpoint
from src.mcp.rate_limit import RateLimiter

PUBLIC_PATHS = ('/health', '/metrics')


def register_observability(server: Any, version: str) -> None:
    """Register ``GET /health`` and ``GET /metrics`` on the server."""

    @server.custom_route('/health', methods=['GET'])
    async def health(request: object) -> Response:
        import time

        return JSONResponse(
            {
                'status': 'ok',
                'server': getattr(server, 'name', 'tianshang-scribe'),
                'version': version,
                'uptime_seconds': round(time.monotonic() - _STARTED, 2),
            }
        )

    @server.custom_route('/metrics', methods=['GET'])
    async def metrics(request: object) -> Response:
        return metrics_endpoint(request)


_STARTED: float = 0.0


def _headers(scope: dict[str, Any]) -> dict[str, str]:
    return {
        key.decode('latin-1').lower(): value.decode('latin-1')
        for key, value in scope.get('headers', [])
    }


class AuthMiddleware:
    """Reject non-bearer HTTP requests when one or more API keys are set."""

    def __init__(self, app: Callable[..., Any], auth_token: str | None) -> None:
        self.app = app
        self.expected: list[str] = [
            part.strip() for part in (auth_token or '').split(',') if part.strip()
        ]

    async def __call__(self, scope: dict[str, Any], receive: Any, send: Any) -> None:
        if self.expected and scope.get('type') == 'http':
            method = scope.get('method', 'GET')
            path = scope.get('path', '')
            if method != 'OPTIONS' and path not in PUBLIC_PATHS:
                header = _headers(scope).get('authorization', '')
                if not header.startswith('Bearer '):
                    await self._reject(scope, receive, send)
                    return
                candidate = header[len('Bearer ') :].strip()
                if not any(
                    hmac.compare_digest(candidate.encode(), expected.encode())
                    for expected in self.expected
                ):
                    await self._reject(scope, receive, send)
                    return
        await self.app(scope, receive, send)

    async def _reject(self, scope: dict[str, Any], receive: Any, send: Any) -> None:
        response = JSONResponse(
            {'error': 'unauthorized', 'message': 'Missing or invalid bearer token'},
            status_code=401,
        )
        await response(scope, receive, send)


class RateLimitMiddleware:
    """Sliding-window rate limiting keyed by client IP / forwarded header."""

    def __init__(self, app: Callable[..., Any], limiter: RateLimiter) -> None:
        self.app = app
        self.limiter = limiter

    async def __call__(self, scope: dict[str, Any], receive: Any, send: Any) -> None:
        if scope.get('type') == 'http':
            method = scope.get('method', 'GET')
            path = scope.get('path', '')
            if (
                method != 'OPTIONS'
                and path not in PUBLIC_PATHS
                and not self.limiter.is_allowed(self._client_id(scope))
            ):
                response = JSONResponse(
                    {
                        'error': 'rate_limited',
                        'message': 'Too many requests; retry later',
                    },
                    status_code=429,
                    headers={'Retry-After': str(self.limiter.window)},
                )
                await response(scope, receive, send)
                return
        await self.app(scope, receive, send)

    def _client_id(self, scope: dict[str, Any]) -> str:
        forwarded = _headers(scope).get('x-forwarded-for', '')
        if forwarded:
            return forwarded.split(',')[0].strip()
        client = scope.get('client')
        return client[0] if client else 'unknown'


def _wrap_app(
    app: Callable[..., Any],
    *,
    auth_token: str | None,
    cors_origins: str | None,
    rate_limiter: RateLimiter,
) -> Callable[..., Any]:
    """Compose CORS, rate limiting and auth around an MCP SDK app."""
    origins = [origin.strip() for origin in (cors_origins or '').split(',') if origin.strip()]
    app = CORSMiddleware(
        app,
        allow_origins=origins or ['*'],
        allow_methods=['*'],
        allow_headers=['*'],
    )
    app = RateLimitMiddleware(app, rate_limiter)
    app = AuthMiddleware(app, auth_token)
    return app


def build_sse_app(
    server: Any,
    *,
    auth_token: str | None,
    cors_origins: str | None,
    rate_limiter: RateLimiter,
    sse_path: str = '/sse',
    message_path: str = '/messages/',
    host: str = '127.0.0.1',
) -> Any:
    """Build the legacy SSE transport app (deprecated upstream, kept for compat)."""
    app = server.sse_app(sse_path=sse_path, message_path=message_path, host=host)
    return _wrap_app(
        app, auth_token=auth_token, cors_origins=cors_origins, rate_limiter=rate_limiter
    )


def build_http_app(
    server: Any,
    *,
    auth_token: str | None,
    cors_origins: str | None,
    rate_limiter: RateLimiter,
    streamable_http_path: str = '/mcp',
    host: str = '127.0.0.1',
) -> Any:
    """Build the Streamable HTTP transport app."""
    app = server.streamable_http_app(streamable_http_path=streamable_http_path, host=host)
    return _wrap_app(
        app, auth_token=auth_token, cors_origins=cors_origins, rate_limiter=rate_limiter
    )


def run_stdio(server: Any) -> None:
    """Run the server over stdio until stdin closes."""
    asyncio.run(server.run_stdio_async())


def run_sse(
    server: Any,
    *,
    host: str,
    port: int,
    auth_token: str | None,
    cors_origins: str | None,
    rate_limiter: RateLimiter,
) -> None:
    """Run the legacy SSE transport server with uvicorn."""
    import uvicorn

    app = build_sse_app(
        server,
        auth_token=auth_token,
        cors_origins=cors_origins,
        rate_limiter=rate_limiter,
        host=host,
    )
    uvicorn.run(app, host=host, port=port, log_level='info')


def run_http(
    server: Any,
    *,
    host: str,
    port: int,
    auth_token: str | None,
    cors_origins: str | None,
    rate_limiter: RateLimiter,
    streamable_http_path: str,
) -> None:
    """Run the Streamable HTTP transport server with uvicorn."""
    import uvicorn

    app = build_http_app(
        server,
        auth_token=auth_token,
        cors_origins=cors_origins,
        rate_limiter=rate_limiter,
        streamable_http_path=streamable_http_path,
        host=host,
    )
    uvicorn.run(app, host=host, port=port, log_level='info')
