"""Centralized runtime configuration for the MCP server.

Unifies the configuration previously scattered across CLI arguments and
environment variables into a single :class:`Settings` model backed by
``pydantic-settings``. Values resolve in priority order:

1. explicit constructor arguments (e.g. CLI overrides);
2. ``TIANSHANG_SCRIBE_*`` environment variables;
3. a local ``.env`` file (if present);
4. the defaults declared below.

A fresh :class:`Settings` instance is cheap to construct (reads the process
environment on each call), so callers may create one per operation instead of
caching — this keeps ``monkeypatch``-based tests and live reloads correct.
"""

from __future__ import annotations

from typing import Literal

from pydantic_settings import BaseSettings, SettingsConfigDict

TransportName = Literal['stdio', 'sse', 'streamable-http']


class Settings(BaseSettings):
    """Server configuration. Environment variables use the ``TIANSHANG_SCRIBE_`` prefix."""

    model_config = SettingsConfigDict(
        env_prefix='TIANSHANG_SCRIBE_',
        env_file='.env',
        env_file_encoding='utf-8',
        extra='ignore',
    )

    transport: TransportName = 'stdio'
    """MCP transport to serve (``stdio``, ``sse`` or ``streamable-http``)."""

    host: str = '127.0.0.1'
    """Bind host for the HTTP transports."""

    port: int = 8080
    """Bind port for the HTTP transports."""

    auth_token: str | None = None
    """Primary bearer token (env: ``TIANSHANG_SCRIBE_AUTH_TOKEN``)."""

    api_keys: str | None = None
    """Comma-separated additional bearer tokens (env: ``TIANSHANG_SCRIBE_API_KEYS``)."""

    cors_origins: str | None = None
    """Comma-separated CORS allowed origins (env: ``TIANSHANG_SCRIBE_CORS_ORIGINS``)."""

    rate_limit_max: int = 100
    """Max requests per client per ``rate_limit_window``."""

    rate_limit_window: int = 60
    """Rate limit window in seconds."""

    mcp_path: str = '/mcp'
    """Streamable HTTP endpoint path."""

    log_level: str = 'INFO'
    """Logging threshold (``DEBUG``/``INFO``/``WARNING``/``ERROR``)."""

    log_json: bool = False
    """Emit structured JSON logs instead of pretty console output."""

    def bearer_tokens(self) -> list[str]:
        """Return the union of configured bearer tokens, de-duplicated."""
        tokens: list[str] = []
        for raw in (self.auth_token, self.api_keys):
            for part in str(raw or '').split(','):
                token = part.strip()
                if token and token not in tokens:
                    tokens.append(token)
        return tokens

    def cors_origin_list(self) -> list[str]:
        """Return the CORS origins as a non-empty list of origin strings."""
        return [origin.strip() for origin in (self.cors_origins or '').split(',') if origin.strip()]


def get_settings() -> Settings:
    """Construct a fresh :class:`Settings` from the current environment."""
    return Settings()
