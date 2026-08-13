"""Unit tests for MCP infrastructure: auth, rate limiting, metrics, prompts, security, server build."""

from __future__ import annotations

import pytest

from src.mcp import auth, prompts, security
from src.mcp.metrics import (
    instrumented,
    metrics_endpoint,
    observe_error,
    observe_operation,
    track_operation,
)
from src.mcp.rate_limit import RateLimiter, RateLimitError
from src.mcp.security import PermissionLevel
from src.mcp.server import INSTRUCTIONS, SERVER_NAME, build_server


class TestAuth:
    def test_open_when_no_keys(self, monkeypatch) -> None:
        monkeypatch.delenv('SCRIBE_AUTH_TOKEN', raising=False)
        monkeypatch.delenv('SCRIBE_API_KEYS', raising=False)
        assert auth.api_key_enabled() is False
        assert auth.validate_api_key('anything') is True
        assert auth.validate_api_key(None) is True

    def test_token_validation(self, monkeypatch) -> None:
        monkeypatch.setenv('SCRIBE_AUTH_TOKEN', 'secret-token')
        assert auth.api_key_enabled() is True
        assert auth.validate_api_key('secret-token') is True
        assert auth.validate_api_key('wrong') is False
        assert auth.validate_api_key(None) is False

    def test_comma_separated_keys(self, monkeypatch) -> None:
        monkeypatch.setenv('SCRIBE_API_KEYS', 'k1, k2,')
        assert auth.validate_api_key('k1') is True
        assert auth.validate_api_key('k2') is True
        assert auth.validate_api_key('k3') is False


class TestRateLimiter:
    def test_validation(self) -> None:
        with pytest.raises(ValueError):
            RateLimiter(max_requests=0)
        with pytest.raises(ValueError):
            RateLimiter(max_requests=2, window_seconds=0)

    def test_allow_up_to_limit(self) -> None:
        limiter = RateLimiter(max_requests=2, window_seconds=60)
        assert limiter.remaining('c1') == 2
        assert limiter.is_allowed('c1') is True
        assert limiter.is_allowed('c1') is True
        assert limiter.is_allowed('c1') is False
        assert limiter.remaining('c1') == 0

    def test_check_raises(self) -> None:
        limiter = RateLimiter(max_requests=1, window_seconds=60)
        limiter.check('c1')
        with pytest.raises(RateLimitError) as exc:
            limiter.check('c1')
        assert exc.value.code == 'rate_limited'
        assert '60s' in str(exc.value)

    def test_independent_clients(self) -> None:
        limiter = RateLimiter(max_requests=1, window_seconds=60)
        assert limiter.is_allowed('a') is True
        assert limiter.is_allowed('b') is True


class TestMetrics:
    def test_track_operation_success(self) -> None:
        with track_operation('test_tool'):
            pass
        observe_operation('test_tool')

    def test_track_operation_error(self) -> None:
        observe_error('test_tool')
        with pytest.raises(ValueError), track_operation('failing_tool'):
            raise ValueError('boom')

    def test_instrumented_wraps_value(self) -> None:
        def add(a: int, b: int) -> int:
            return a + b

        wrapped = instrumented(add, 'add_tool')
        assert wrapped(1, 2) == 3

    def test_metrics_endpoint(self) -> None:
        resp = metrics_endpoint(object())
        assert resp.media_type.startswith('text/plain')
        body = resp.body.decode()
        assert 'scribe_operation_duration_seconds' in body


class TestPrompts:
    def test_register_prompts(self) -> None:
        import asyncio

        from mcp.server.mcpserver import MCPServer

        server = MCPServer(name='x', version='1')
        prompts.register_prompts(server)
        listed = asyncio.run(server.list_prompts())
        names = {p.name for p in listed}
        assert {'generate_report', 'batch_fill_templates', 'convert_and_archive'} <= names

    def test_prompt_texts(self) -> None:
        assert 'create_office_document' in prompts._generate_report('T', 'S1,S2')
        assert 'fill_template' in prompts._batch_fill_templates('/t.docx', '/d.csv')
        assert 'convert_document' in prompts._convert_and_archive('*.docx', 'C')
        assert 'extract_document_data' in prompts._extract_and_analyze('/doc.docx')
        assert 'format="pptx"' in prompts._create_presentation('T', 'Intro: text')


class TestSecurity:
    def test_levels(self) -> None:
        assert security.levels_for('extract_document_data') == {PermissionLevel.READ_ONLY}
        assert security.levels_for('edit_office_document') == {PermissionLevel.DESTRUCTIVE}
        assert security.levels_for('unknown_tool') == {PermissionLevel.STANDARD}

    def test_is_read_only(self) -> None:
        assert security.is_read_only('extract_document_data') is True
        assert security.is_read_only('create_office_document') is False

    def test_is_destructive(self) -> None:
        assert security.is_destructive('edit_office_document') is True
        assert security.is_destructive('fill_template') is False

    def test_is_idempotent(self) -> None:
        assert security.is_idempotent('compare_documents') is True
        assert security.is_idempotent('create_office_document') is False

    def test_idempotent_table_covers_registry(self) -> None:
        from src.mcp.tools._registry import TOOLS

        all_names = {t['name'] for t in TOOLS}
        extra = set(security.IDEMPOTENT_TOOLS) - all_names
        assert not extra

    def test_idempotent_equals_read_only(self) -> None:
        from src.mcp.tools._registry import TOOLS

        for entry in TOOLS:
            name = entry['name']
            assert security.is_idempotent(name) == security.is_read_only(name), name

    def test_check_permission(self) -> None:
        assert security.check_permission('create_office_document', {'create_office_document'}) is True
        assert security.check_permission('extract_document_data', set()) is True
        assert security.check_permission('create_office_document', set()) is False


class TestBuildServer:
    def test_build_server(self) -> None:
        server = build_server(version='9.9.9')
        assert getattr(server, 'name', SERVER_NAME) == SERVER_NAME

    def test_instructions_embedded(self) -> None:
        assert 'edit_office_document' in INSTRUCTIONS
        assert 'Authorization' in INSTRUCTIONS
