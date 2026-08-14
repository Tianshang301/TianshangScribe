"""HTTP transport wiring for the TianshangScribe MCP Server.

Builds on the official MCP SDK apps (:meth:`mcp.server.mcpserver.MCPServer.sse_app`
and ``.streamable_http_app``) and wraps them with CORS, bearer-token auth,
role-based access control and sliding-window rate limiting. ``/health`` and
``/metrics`` are registered on the server itself via ``custom_route`` so they
stay available on every transport.

Middleware is implemented as pure ASGI classes (not ``BaseHTTPMiddleware``) so
the long-lived SSE/streamable-HTTP GET streams are never buffered.
"""

from __future__ import annotations

import asyncio
import hmac
import json
from collections.abc import Callable
from typing import Any

from starlette.middleware.cors import CORSMiddleware
from starlette.responses import JSONResponse, Response

from tianshang_scribe.mcp.metrics import metrics_endpoint
from tianshang_scribe.mcp.rate_limit import RateLimiter
from tianshang_scribe.mcp.security import parse_role, role_allows
from tianshang_scribe.utils.config import Settings
from tianshang_scribe.utils.logging import get_logger

PUBLIC_PATHS = ('/health', '/metrics')
ROLE_HEADER = 'x-scribe-role'


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


async def _read_body(receive: Any) -> bytes:
    """Consume the full request body from the ASGI ``receive`` callable."""
    chunks: list[bytes] = []
    while True:
        message = await receive()
        if message.get('type') != 'http.request':
            break
        chunks.append(message.get('body', b''))
        if not message.get('more_body'):
            break
    return b''.join(chunks)


def _replaying_receive(body: bytes) -> Callable[..., Any]:
    """Return a ``receive`` that replays an already-consumed body once."""

    async def receive() -> dict[str, Any]:
        return {'type': 'http.request', 'body': body, 'more_body': False}

    return receive


def _parse_rpc_tool(body: bytes) -> tuple[str, str]:
    """Extract (method, tool name) from a JSON-RPC body.

    Returns (``''``, ``''``) when the body is not JSON-RPC or the method is not
    ``tools/call``. A batch body checks every request until one names a tool.
    """
    try:
        payload = json.loads(body.decode('utf-8'))
    except (json.JSONDecodeError, UnicodeDecodeError):
        return '', ''
    items = payload if isinstance(payload, list) else [payload]
    for item in items:
        if not isinstance(item, dict):
            continue
        method = item.get('method', '')
        if method == 'tools/call':
            params = item.get('params', {})
            if isinstance(params, dict):
                name = params.get('name')
                return method, str(name) if name else ''
    return '', ''


class AuthMiddleware:
    """Reject non-bearer HTTP requests when one or more API keys are set.

    Distinguishes unauthenticated (no ``Authorization`` header, ``401``) from
    unauthorized (header present but key invalid, ``403``). Requests to
    :data:`PUBLIC_PATHS` (``/health``, ``/metrics``) and CORS ``OPTIONS``
    preflights always pass through.
    """

    def __init__(self, app: Callable[..., Any], auth_token: str | None) -> None:
        """Store the ASGI app and parse the allowed bearer-token list."""
        self.app = app
        self.expected: list[str] = [
            part.strip() for part in (auth_token or '').split(',') if part.strip()
        ]

    async def __call__(self, scope: dict[str, Any], receive: Any, send: Any) -> None:
        """Enforce bearer-token auth on HTTP requests before passing through."""
        if self.expected and scope.get('type') == 'http':
            method = scope.get('method', 'GET')
            path = scope.get('path', '')
            if method != 'OPTIONS' and path not in PUBLIC_PATHS:
                header = _headers(scope).get('authorization', '')
                if not header.startswith('Bearer '):
                    get_logger('scribe.http').warning(
                        'auth_rejected',
                        method=method,
                        path=path,
                        status_code=401,
                        reason='missing_token',
                    )
                    await self._reject(scope, receive, send, status_code=401)
                    return
                candidate = header[len('Bearer ') :].strip()
                if not any(
                    hmac.compare_digest(candidate.encode(), expected.encode())
                    for expected in self.expected
                ):
                    get_logger('scribe.http').warning(
                        'auth_rejected',
                        method=method,
                        path=path,
                        status_code=403,
                        reason='invalid_token',
                    )
                    await self._reject(scope, receive, send, status_code=403)
                    return
        await self.app(scope, receive, send)

    async def _reject(
        self, scope: dict[str, Any], receive: Any, send: Any, *, status_code: int
    ) -> None:
        response = JSONResponse(
            {
                'error': 'unauthorized',
                'message': (
                    'Missing bearer token' if status_code == 401 else 'Invalid bearer token'
                ),
            },
            status_code=status_code,
        )
        await response(scope, receive, send)


class RbacMiddleware:
    """Reject JSON-RPC ``tools/call`` requests for tools the role cannot use.

    The client declares its role with the ``X-Scribe-Role`` header (values:
    ``viewer`` / ``editor`` / ``owner``). When the header is absent or unknown
    the most privileged role (``owner``) is assumed so existing deployments
    keep working unchanged. Requests are allowed when the caller's role may
    invoke the requested tool per :data:`tianshang_scribe.mcp.security.ROLE_TOOL_MATRIX`.
    """

    def __init__(self, app: Callable[..., Any]) -> None:
        """Store the ASGI app wrapped by the RBAC check."""
        self.app = app

    async def __call__(self, scope: dict[str, Any], receive: Any, send: Any) -> None:
        """Enforce the role→tool matrix on JSON-RPC tool calls."""
        if scope.get('type') == 'http' and scope.get('method') == 'POST':
            body = await _read_body(receive)
            method, tool_name = _parse_rpc_tool(body)
            if method == 'tools/call' and tool_name:
                role = parse_role(_headers(scope).get(ROLE_HEADER))
                if not role_allows(role, tool_name):
                    get_logger('scribe.http').warning(
                        'rbac_rejected',
                        method=method,
                        tool=tool_name,
                        role=role.value,
                        status_code=403,
                    )
                    response = JSONResponse(
                        {
                            'error': 'forbidden',
                            'message': f'Role {role.value!r} is not allowed to call {tool_name!r}.',
                        },
                        status_code=403,
                    )
                    await response(scope, receive, send)
                    return
            await self.app(scope, _replaying_receive(body), send)
            return
        await self.app(scope, receive, send)


class RateLimitMiddleware:
    """Sliding-window rate limiting keyed by client IP / forwarded header."""

    def __init__(self, app: Callable[..., Any], limiter: RateLimiter) -> None:
        """Store the ASGI app and the rate limiter instance."""
        self.app = app
        self.limiter = limiter

    async def __call__(self, scope: dict[str, Any], receive: Any, send: Any) -> None:
        """Reject HTTP requests that exceed the rate limit before passing through."""
        if scope.get('type') == 'http':
            method = scope.get('method', 'GET')
            path = scope.get('path', '')
            if (
                method != 'OPTIONS'
                and path not in PUBLIC_PATHS
                and not self.limiter.is_allowed(self._client_id(scope))
            ):
                client_id = self._client_id(scope)
                get_logger('scribe.http').warning(
                    'rate_limited',
                    method=method,
                    path=path,
                    status_code=429,
                    client_id=client_id,
                )
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
    """Compose CORS, rate limiting, auth and RBAC around an MCP SDK app."""
    origins = [origin.strip() for origin in (cors_origins or '').split(',') if origin.strip()]
    app = CORSMiddleware(
        app,
        allow_origins=origins or ['*'],
        allow_methods=['*'],
        allow_headers=['*'],
    )
    app = RateLimitMiddleware(app, rate_limiter)
    app = AuthMiddleware(app, auth_token)
    app = RbacMiddleware(app)
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


def _uvicorn_log_config() -> dict[str, Any]:
    """Return a uvicorn log config aligned with the structlog pipeline.

    ``level`` is read from the current :class:`Settings` so ``SCRIBE_LOG_LEVEL``
    and ``SCRIBE_LOG_JSON`` control uvicorn's access/error logs too.
    """
    settings = Settings()
    level = settings.log_level.upper()
    formatter = 'json' if settings.log_json else 'default'
    return {
        'version': 1,
        'disable_existing_loggers': False,
        'formatters': {
            'default': {'format': '%(asctime)s %(levelname)s %(name)s: %(message)s'},
            'json': {
                'format': '{"ts":"%(asctime)s","level":"%(levelname)s","logger":"%(name)s","message":"%(message)s"}',
            },
        },
        'handlers': {
            'default': {'class': 'logging.StreamHandler', 'formatter': formatter},
        },
        'loggers': {
            'uvicorn': {'handlers': ['default'], 'level': level, 'propagate': False},
            'uvicorn.access': {'handlers': ['default'], 'level': level, 'propagate': False},
            'uvicorn.error': {'handlers': ['default'], 'level': level, 'propagate': False},
        },
    }


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
    uvicorn.run(app, host=host, port=port, log_config=_uvicorn_log_config())


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
    uvicorn.run(app, host=host, port=port, log_config=_uvicorn_log_config())
