"""Unit tests for centralized configuration (tianshang_scribe/utils/config.py)."""

from __future__ import annotations

import pytest
from pydantic import ValidationError

from tianshang_scribe.utils.config import Settings


class TestSettingsDefaults:
    def test_defaults(self) -> None:
        settings = Settings()
        assert settings.transport == 'stdio'
        assert settings.host == '127.0.0.1'
        assert settings.port == 8080
        assert settings.auth_token is None
        assert settings.api_keys is None
        assert settings.cors_origins is None
        assert settings.rate_limit_max == 100
        assert settings.rate_limit_window == 60
        assert settings.mcp_path == '/mcp'

    def test_no_tokens_by_default(self) -> None:
        assert Settings().bearer_tokens() == []

    def test_empty_cors_list(self) -> None:
        assert Settings().cors_origin_list() == []


class TestSettingsEnvOverrides:
    def test_auth_token_env(self, monkeypatch) -> None:
        monkeypatch.setenv('SCRIBE_AUTH_TOKEN', 'secret')
        assert Settings().bearer_tokens() == ['secret']

    def test_api_keys_env(self, monkeypatch) -> None:
        monkeypatch.setenv('SCRIBE_API_KEYS', 'k1, k2,')
        assert Settings().bearer_tokens() == ['k1', 'k2']

    def test_both_token_sources_merged_deduplicated(self, monkeypatch) -> None:
        monkeypatch.setenv('SCRIBE_AUTH_TOKEN', 'primary')
        monkeypatch.setenv('SCRIBE_API_KEYS', 'primary, extra')
        assert Settings().bearer_tokens() == ['primary', 'extra']

    def test_host_port_env(self, monkeypatch) -> None:
        monkeypatch.setenv('SCRIBE_HOST', '192.168.1.5')
        monkeypatch.setenv('SCRIBE_PORT', '9000')
        settings = Settings()
        assert settings.host == '192.168.1.5'
        assert settings.port == 9000

    def test_cors_origins_env(self, monkeypatch) -> None:
        monkeypatch.setenv('SCRIBE_CORS_ORIGINS', 'https://a.dev, https://b.dev')
        assert Settings().cors_origin_list() == ['https://a.dev', 'https://b.dev']

    def test_rate_limit_env(self, monkeypatch) -> None:
        monkeypatch.setenv('SCRIBE_RATE_LIMIT_MAX', '50')
        monkeypatch.setenv('SCRIBE_RATE_LIMIT_WINDOW', '30')
        settings = Settings()
        assert settings.rate_limit_max == 50
        assert settings.rate_limit_window == 30

    def test_transport_env(self, monkeypatch) -> None:
        monkeypatch.setenv('SCRIBE_TRANSPORT', 'sse')
        assert Settings().transport == 'sse'

    def test_unrelated_env_ignored(self, monkeypatch) -> None:
        monkeypatch.setenv('SCRIBE_SOMETHING_ELSE', 'x')
        assert Settings().auth_token is None


class TestSettingsValidation:
    def test_invalid_transport_rejected(self, monkeypatch) -> None:
        monkeypatch.setenv('SCRIBE_TRANSPORT', 'carrier-pigeon')
        with pytest.raises(ValidationError):
            Settings()

    def test_invalid_port_rejected(self, monkeypatch) -> None:
        monkeypatch.setenv('SCRIBE_PORT', 'not-a-number')
        with pytest.raises(ValidationError):
            Settings()

    def test_constructor_args_override_env(self, monkeypatch) -> None:
        monkeypatch.setenv('SCRIBE_HOST', '192.168.1.5')
        assert Settings(host='10.0.0.1').host == '10.0.0.1'


class TestGetSettings:
    def test_returns_fresh_settings(self, monkeypatch) -> None:
        from tianshang_scribe.utils.config import get_settings

        assert get_settings().host == '127.0.0.1'
        monkeypatch.setenv('SCRIBE_HOST', '192.168.1.5')
        assert get_settings().host == '192.168.1.5'
