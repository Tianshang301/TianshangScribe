"""API-key authentication for the HTTP transports.

A single ``SCRIBE_AUTH_TOKEN`` (or a comma-separated ``SCRIBE_API_KEYS`` list)
guards the SSE and streamable-HTTP endpoints. stdio mode is unaffected. When
no key is configured the HTTP endpoint is open (local development).
"""

from __future__ import annotations

import hmac
import os


def _configured_keys() -> list[str]:
    """Return the list of valid API keys from the environment."""
    keys: list[str] = []
    for raw in (os.environ.get('SCRIBE_AUTH_TOKEN'), os.environ.get('SCRIBE_API_KEYS')):
        for part in str(raw or '').split(','):
            token = part.strip()
            if token:
                keys.append(token)
    return keys


def api_key_enabled() -> bool:
    """Return whether API-key authentication is configured."""
    return bool(_configured_keys())


def validate_api_key(token: str | None) -> bool:
    """Return whether ``token`` is a valid API key.

    When no keys are configured the endpoint is open and any request is
    allowed (returns ``True``).
    """
    if not api_key_enabled():
        return True
    if not token:
        return False
    candidate = token.strip()
    return any(hmac.compare_digest(candidate.encode(), key.encode()) for key in _configured_keys())
