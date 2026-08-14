"""Persistent stdio MCP client shared by the smoke-test scripts.

The MCP SDK requires ``initialize`` before any other request, and responses
are delivered over the same stdio stream, so these scripts use one persistent
subprocess per session instead of the old one-subprocess-per-call pattern.
"""

from __future__ import annotations

import contextlib
import json
import subprocess
import sys
import time
from pathlib import Path
from typing import cast

PROJECT_ROOT = str(Path(__file__).resolve().parent.parent.parent)


class MCPClient:
    """Tiny MCP client over a long-lived stdio subprocess."""

    def __init__(self, timeout: float = 30.0) -> None:
        self.timeout = timeout
        self.proc = subprocess.Popen(
            [sys.executable, '-m', 'tianshang_scribe.mcp.server'],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding='utf-8',
            cwd=PROJECT_ROOT,
        )
        self._id = 0

    def initialize(self) -> dict:
        result = self.call(
            'initialize',
            {
                'protocolVersion': '2024-11-05',
                'capabilities': {},
                'clientInfo': {'name': 'mcp-client', 'version': '1.0'},
            },
        )
        self.notify('notifications/initialized')
        return result

    def _send(self, message: dict) -> None:
        stdin = self.proc.stdin
        assert stdin is not None
        stdin.write(json.dumps(message) + '\n')
        stdin.flush()

    def _read(self) -> dict:
        stdout = self.proc.stdout
        assert stdout is not None
        deadline = time.time() + self.timeout
        while time.time() < deadline:
            line = stdout.readline()
            if not line:
                raise RuntimeError(f'MCP server closed stdout; stderr: {self.stderr()}')
            return cast(dict, json.loads(line))
        raise TimeoutError('Timed out waiting for MCP server response')

    def stderr(self) -> str:
        if self.proc.stderr:
            return cast(str, self.proc.stderr.read())
        return ''

    def notify(self, method: str, params: dict | None = None) -> None:
        self._send({'jsonrpc': '2.0', 'method': method, 'params': params or {}})

    def call(self, method: str, params: dict | None = None) -> dict:
        self._id += 1
        msg_id = self._id
        self._send({'jsonrpc': '2.0', 'id': msg_id, 'method': method, 'params': params or {}})
        resp = self._read()
        while resp.get('id') != msg_id:
            resp = self._read()
        return resp

    def tool(self, name: str, arguments: dict) -> dict:
        resp = self.call('tools/call', {'name': name, 'arguments': arguments})
        if 'error' in resp:
            raise RuntimeError(f'tools/call {name} error: {resp["error"]}')
        return cast(dict, json.loads(resp['result']['content'][0]['text']))

    def close(self) -> None:
        with contextlib.suppress(Exception):
            stdin = self.proc.stdin
            if stdin is not None:
                stdin.close()
        try:
            self.proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            self.proc.kill()
            self.proc.wait()
