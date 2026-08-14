"""Unit tests for structlog integration (src/utils/logging.py)."""

from __future__ import annotations

import json
import re

import pytest
import structlog

from src.utils.config import Settings
from src.utils.logging import configure_logging, get_logger, log_http_event

_ANSI = re.compile(r'\x1b\[[0-9;]*m')


def _strip_ansi(text: str) -> str:
    return _ANSI.sub('', text)


@pytest.fixture(autouse=True)
def _reset_structlog() -> None:
    structlog.reset_defaults()
    yield
    structlog.reset_defaults()


def _capture(capsys: pytest.CaptureFixture[str], event: str, **kw: object) -> str:
    get_logger('scribe.test').info(event, **kw)
    return capsys.readouterr().out


class TestConsoleMode:
    def test_console_output(self, capsys) -> None:
        configure_logging(Settings(log_json=False))
        out = _strip_ansi(_capture(capsys, 'ping', key='v'))
        assert 'ping' in out
        assert 'key=v' in out

    def test_info_level_filters_debug(self, capsys) -> None:
        configure_logging(Settings(log_level='INFO', log_json=False))
        get_logger('scribe.test').debug('hidden')
        assert capsys.readouterr().out == ''


class TestJsonMode:
    def test_json_output_is_parseable(self, capsys) -> None:
        configure_logging(Settings(log_json=True))
        out = _capture(capsys, 'ping', key='v')
        payload = json.loads(out.strip())
        assert payload['event'] == 'ping'
        assert payload['key'] == 'v'
        assert payload['level'] == 'info'

    def test_json_has_timestamp_and_level(self, capsys) -> None:
        configure_logging(Settings(log_json=True))
        payload = json.loads(_capture(capsys, 'evt').strip())
        assert 'timestamp' in payload
        assert 'level' in payload


class TestConfigSources:
    def test_env_log_json(self, monkeypatch, capsys) -> None:
        monkeypatch.setenv('SCRIBE_LOG_JSON', 'true')
        configure_logging()
        out = _capture(capsys, 'env_event')
        assert json.loads(out.strip())['event'] == 'env_event'

    def test_env_log_level(self, monkeypatch, capsys) -> None:
        monkeypatch.setenv('SCRIBE_LOG_LEVEL', 'WARNING')
        configure_logging()
        get_logger('scribe.test').info('filtered')
        assert capsys.readouterr().out == ''
        get_logger('scribe.test').warning('kept')
        assert 'kept' in capsys.readouterr().out

    def test_settings_log_fields(self, monkeypatch) -> None:
        monkeypatch.setenv('SCRIBE_LOG_JSON', '1')
        monkeypatch.setenv('SCRIBE_LOG_LEVEL', 'ERROR')
        settings = Settings()
        assert settings.log_json is True
        assert settings.log_level == 'ERROR'


class TestLogHttpEvent:
    def test_log_http_event_json(self, capsys) -> None:
        configure_logging(Settings(log_json=True))
        log_http_event(
            'auth_rejected', method='POST', path='/mcp', status_code=401, reason='missing'
        )
        payload = json.loads(capsys.readouterr().out.strip())
        assert payload['event'] == 'auth_rejected'
        assert payload['method'] == 'POST'
        assert payload['status_code'] == 401
        assert payload['reason'] == 'missing'

    def test_log_http_event_console(self, capsys) -> None:
        configure_logging(Settings(log_json=False))
        log_http_event('rate_limited', method='GET', path='/sse')
        assert 'rate_limited' in capsys.readouterr().out

    def test_get_logger_returns_bound_logger(self) -> None:
        assert callable(get_logger('x').info)
