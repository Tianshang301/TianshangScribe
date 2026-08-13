"""Unit tests for MCP HTTP transport middleware (auth, rate limiting, CORS)."""

from __future__ import annotations

import asyncio
import json
from typing import Any

from src.mcp.rate_limit import RateLimiter
from src.mcp.transport import (
    PUBLIC_PATHS,
    AuthMiddleware,
    RateLimitMiddleware,
    build_http_app,
    build_sse_app,
)


class RecordingApp:
    """ASGI app stub that records that it was reached."""

    def __init__(self) -> None:
        self.called = False

    async def __call__(self, scope: dict[str, Any], receive: Any, send: Any) -> None:
        self.called = True


class Scope:
    """Minimal ASGI scope builder."""

    @staticmethod
    def http(path: str = '/mcp', method: str = 'POST', auth: str | None = None) -> dict[str, Any]:
        headers: list[tuple[bytes, bytes]] = []
        if auth:
            headers.append((b'authorization', auth.encode()))
        return {
            'type': 'http',
            'method': method,
            'path': path,
            'headers': headers,
            'client': ('127.0.0.1', 5555),
        }


def run(coro: Any) -> Any:
    return asyncio.run(coro)


async def invoke(app: Any, scope: dict[str, Any]) -> tuple[int, str]:
    """Drive an ASGI app and capture the response status and body."""
    sent: list[dict[str, Any]] = []
    received: dict[str, Any] = {}

    async def receive() -> dict[str, Any]:
        return received

    async def send(message: dict[str, Any]) -> None:
        sent.append(message)

    await app(scope, receive, send)
    status = next((m.get('status') for m in sent if m.get('type') == 'http.response.start'), 200)
    body = b''.join(m.get('body', b'') for m in sent if m.get('type') == 'http.response.body')
    try:
        return int(status), body.decode('utf-8', 'replace')
    except (TypeError, ValueError):
        return 200, body.decode('utf-8', 'replace')


class TestPublicPaths:
    def test_health_and_metrics_exempt(self) -> None:
        assert PUBLIC_PATHS == ('/health', '/metrics')


class TestAuthMiddleware:
    def test_open_when_no_keys(self) -> None:
        inner = RecordingApp()
        app = AuthMiddleware(inner, None)
        status, _ = run(invoke(app, Scope.http()))
        assert status == 200
        assert inner.called is True

    def test_missing_header_returns_401(self) -> None:
        inner = RecordingApp()
        app = AuthMiddleware(inner, 'secret')
        status, body = run(invoke(app, Scope.http()))
        assert status == 401
        assert 'Missing bearer token' in body
        assert inner.called is False

    def test_invalid_token_returns_403(self) -> None:
        inner = RecordingApp()
        app = AuthMiddleware(inner, 'secret')
        status, body = run(invoke(app, Scope.http(auth='Bearer wrong')))
        assert status == 403
        assert 'Invalid bearer token' in body
        assert inner.called is False

    def test_valid_token_passes(self) -> None:
        inner = RecordingApp()
        app = AuthMiddleware(inner, 'secret')
        status, _ = run(invoke(app, Scope.http(auth='Bearer secret')))
        assert status == 200
        assert inner.called is True

    def test_multiple_keys_accepted(self) -> None:
        inner = RecordingApp()
        app = AuthMiddleware(inner, 'k1, k2,')
        assert run(invoke(app, Scope.http(auth='Bearer k2')))[0] == 200
        assert inner.called is True

    def test_non_bearer_header_returns_401(self) -> None:
        inner = RecordingApp()
        app = AuthMiddleware(inner, 'secret')
        status, _ = run(invoke(app, Scope.http(auth='Basic abc')))
        assert status == 401
        assert inner.called is False

    def test_health_exempt_from_auth(self) -> None:
        inner = RecordingApp()
        app = AuthMiddleware(inner, 'secret')
        status, _ = run(invoke(app, Scope.http(path='/health', method='GET')))
        assert status == 200
        assert inner.called is True

    def test_metrics_exempt_from_auth(self) -> None:
        inner = RecordingApp()
        app = AuthMiddleware(inner, 'secret')
        status, _ = run(invoke(app, Scope.http(path='/metrics', method='GET')))
        assert status == 200
        assert inner.called is True

    def test_options_preflight_exempt(self) -> None:
        inner = RecordingApp()
        app = AuthMiddleware(inner, 'secret')
        status, _ = run(invoke(app, Scope.http(method='OPTIONS')))
        assert status == 200
        assert inner.called is True

    def test_whitespace_trimmed_candidate(self) -> None:
        inner = RecordingApp()
        app = AuthMiddleware(inner, 'secret')
        status, _ = run(invoke(app, Scope.http(auth='Bearer  secret ')))
        assert status == 200
        assert inner.called is True


class TestRateLimitMiddleware:
    def test_allowed_within_limit(self) -> None:
        inner = RecordingApp()
        limiter = RateLimiter(max_requests=2, window_seconds=60)
        app = RateLimitMiddleware(inner, limiter)
        for _ in range(2):
            status, _ = run(invoke(app, Scope.http()))
            assert status == 200
        assert inner.called is True

    def test_exceeds_limit_returns_429(self) -> None:
        inner = RecordingApp()
        limiter = RateLimiter(max_requests=1, window_seconds=60)
        app = RateLimitMiddleware(inner, limiter)
        assert run(invoke(app, Scope.http()))[0] == 200
        status, body = run(invoke(app, Scope.http()))
        assert status == 429
        assert 'rate_limited' in body

    def test_health_exempt_from_rate_limit(self) -> None:
        inner = RecordingApp()
        limiter = RateLimiter(max_requests=1, window_seconds=60)
        app = RateLimitMiddleware(inner, limiter)
        assert run(invoke(app, Scope.http(path='/health', method='GET')))[0] == 200
        assert run(invoke(app, Scope.http(path='/health', method='GET')))[0] == 200

    def test_metrics_exempt_from_rate_limit(self) -> None:
        inner = RecordingApp()
        limiter = RateLimiter(max_requests=1, window_seconds=60)
        app = RateLimitMiddleware(inner, limiter)
        assert run(invoke(app, Scope.http(path='/metrics', method='GET')))[0] == 200
        assert run(invoke(app, Scope.http(path='/metrics', method='GET')))[0] == 200

    def test_options_exempt_from_rate_limit(self) -> None:
        inner = RecordingApp()
        limiter = RateLimiter(max_requests=1, window_seconds=60)
        app = RateLimitMiddleware(inner, limiter)
        assert run(invoke(app, Scope.http(method='OPTIONS')))[0] == 200
        assert run(invoke(app, Scope.http(method='OPTIONS')))[0] == 200

    def test_forwarded_for_used_as_client_id(self) -> None:
        inner = RecordingApp()
        limiter = RateLimiter(max_requests=1, window_seconds=60)
        app = RateLimitMiddleware(inner, limiter)
        scope = Scope.http()
        scope['headers'].append((b'x-forwarded-for', b'203.0.113.9'))
        assert run(invoke(app, scope))[0] == 200
        assert run(invoke(app, scope))[0] == 429


class TestWrapApps:
    def test_build_sse_app_requires_auth(self) -> None:
        server = _FakeMCP('sse')
        app = build_sse_app(
            server,
            auth_token='tok',
            cors_origins=None,
            rate_limiter=RateLimiter(max_requests=10, window_seconds=60),
        )
        status, _ = run(invoke(app, Scope.http(path='/sse', method='GET')))
        assert status == 401
        status, _ = run(invoke(app, Scope.http(path='/sse', method='GET', auth='Bearer tok')))
        assert status == 200

    def test_build_http_app_requires_auth(self) -> None:
        server = _FakeMCP('http')
        app = build_http_app(
            server,
            auth_token='tok',
            cors_origins=None,
            rate_limiter=RateLimiter(max_requests=10, window_seconds=60),
        )
        status, _ = run(invoke(app, Scope.http(path='/mcp', method='POST')))
        assert status == 401
        status, _ = run(invoke(app, Scope.http(path='/mcp', method='POST', auth='Bearer tok')))
        assert status == 200


class _FakeMCP:
    """Minimal stand-in for the MCP SDK app produced by ``build_sse_app``."""

    def __init__(self, kind: str) -> None:
        self.kind = kind

    def sse_app(self, **kwargs: Any) -> RecordingApp:
        return RecordingApp()

    def streamable_http_app(self, **kwargs: Any) -> RecordingApp:
        return RecordingApp()


def test_json_payload_structure() -> None:
    inner = RecordingApp()
    app = AuthMiddleware(inner, 'secret')
    _, body = run(invoke(app, Scope.http()))
    payload = json.loads(body)
    assert payload['error'] == 'unauthorized'
    assert payload['message']
