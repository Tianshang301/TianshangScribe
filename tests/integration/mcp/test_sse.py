"""Tests for the SSE MCP Server transport (official SDK protocol).

The SDK SSE transport returns ``202 Accepted`` on POST /message and delivers
JSON-RPC responses as ``event: message`` frames on the /sse stream, so the
tests implement a small SSE client instead of reading the POST body.
"""

import http.client
import json
import queue
import subprocess
import sys
import threading
import time
from pathlib import Path
from typing import cast

PROJECT_ROOT = str(Path(__file__).resolve().parent.parent.parent)

INITIALIZE = {
    'jsonrpc': '2.0',
    'id': 1,
    'method': 'initialize',
    'params': {
        'protocolVersion': '2024-11-05',
        'capabilities': {},
        'clientInfo': {'name': 'test-sse', 'version': '1.0'},
    },
}


def _wait_for_server(host: str, port: int, timeout: float = 10.0) -> None:
    """Poll until the server is accepting connections."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            conn = http.client.HTTPConnection(host, port, timeout=1)
            conn.request('GET', '/sse')
            conn.getresponse()
            conn.close()
            time.sleep(0.2)
            return
        except (ConnectionRefusedError, OSError):
            time.sleep(0.3)
    raise AssertionError(f'Server on {host}:{port} did not start within {timeout}s')


def _start_server(port: int) -> subprocess.Popen:
    proc = subprocess.Popen(
        [
            sys.executable,
            '-m',
            'tianshang_scribe.mcp.server',
            '--transport',
            'sse',
            '--port',
            str(port),
        ],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        cwd=PROJECT_ROOT,
    )
    _wait_for_server('127.0.0.1', port)
    return proc


def _stop_server(proc: subprocess.Popen) -> None:
    proc.terminate()
    try:
        proc.wait(timeout=5)
    except subprocess.TimeoutExpired:
        proc.kill()
        proc.wait()
    if proc.stdout:
        proc.stdout.close()
    if proc.stderr:
        proc.stderr.close()


class SseClient:
    """Minimal MCP SSE client reading responses off the /sse stream."""

    def __init__(self, host: str, port: int, timeout: float = 15.0) -> None:
        self.host = host
        self.port = port
        self._timeout = timeout
        self._events: queue.Queue[tuple[str, str]] = queue.Queue()
        self._conn = http.client.HTTPConnection(host, port, timeout=timeout)
        self._conn.request('GET', '/sse')
        self._resp = self._conn.getresponse()
        if self._resp.status != 200:
            raise AssertionError(f'GET /sse returned {self._resp.status}')
        self.message_path, self.session_id = self._read_endpoint()
        self._reader = threading.Thread(target=self._read_loop, daemon=True)
        self._reader.start()

    def _read_endpoint(self) -> tuple[str, str]:
        """Read frames until the endpoint event and extract the message path."""
        event: str | None = None
        data = ''
        deadline = time.time() + self._timeout
        while time.time() < deadline:
            line = self._resp.readline()
            if not line:
                break
            text = line.decode('utf-8', 'replace').rstrip('\r\n')
            if not text:
                if event == 'endpoint':
                    path, _, query = data.partition('?')
                    session = ''
                    for part in query.split('&'):
                        if part.startswith('session_id='):
                            session = part[len('session_id=') :].strip()
                            break
                    return path, session
                event, data = None, ''
                continue
            if text.startswith('event: '):
                event = text[len('event: ') :].strip()
            elif text.startswith('data: '):
                data = text[len('data: ') :].strip()
        raise AssertionError('No endpoint event received')

    def _read_loop(self) -> None:
        """Parse SSE frames after the endpoint event into the queue."""
        event: str | None = None
        data: list[str] = []
        while True:
            try:
                line = self._resp.readline()
            except Exception:
                break
            if not line:
                break
            text = line.decode('utf-8', 'replace').rstrip('\r\n')
            if not text:
                if event:
                    self._events.put((event, '\n'.join(data)))
                event, data = None, []
                continue
            if text.startswith('event: '):
                event = text[len('event: ') :].strip()
            elif text.startswith('data: '):
                data.append(text[len('data: ') :].strip())

    def request(self, method: str, params: dict, msg_id: int) -> dict:
        """POST a JSON-RPC request and wait for its matching SSE response."""
        body = json.dumps(
            {
                'jsonrpc': '2.0',
                'id': msg_id,
                'method': method,
                'params': params,
            }
        )
        conn = http.client.HTTPConnection(self.host, self.port, timeout=10)
        conn.request(
            'POST',
            f'{self.message_path}?session_id={self.session_id}',
            body=body,
            headers={'Content-Type': 'application/json'},
        )
        resp = conn.getresponse()
        status = resp.status
        resp.read()
        conn.close()
        if status != 202:
            raise AssertionError(f'POST {method} returned {status}')

        deadline = time.time() + self._timeout
        while time.time() < deadline:
            try:
                event, raw = self._events.get(timeout=self._timeout)
            except queue.Empty:
                break
            if event != 'message':
                continue
            message = json.loads(raw)
            if message.get('id') == msg_id:
                return cast(dict, message)
        raise AssertionError(f'No response received for {method} (id={msg_id})')

    def close(self) -> None:
        self._conn.close()


def test_sse_server_lifecycle():
    """SSE server handles initialize + tools/list + tools/call."""
    port = 18765
    proc = _start_server(port)
    client = None
    try:
        client = SseClient('127.0.0.1', port)

        init = client.request('initialize', INITIALIZE['params'], 1)
        assert init['result']['serverInfo']['name'] == 'tianshang-scribe'

        tools = client.request('tools/list', {}, 2)
        assert len(tools['result']['tools']) == 12

        call = client.request(
            'tools/call',
            {
                'name': 'create_office_document',
                'arguments': {
                    'format': 'docx',
                    'content': [{'type': 'heading', 'text': 'Test', 'level': 1}],
                },
            },
            3,
        )
        content = call['result']['content'][0]['text']
        result = json.loads(content)
        assert result['success']
    finally:
        if client:
            client.close()
        _stop_server(proc)


def test_missing_session_rejected():
    """POST /message with invalid session_id returns 400."""
    port = 18766
    proc = _start_server(port)
    try:
        conn = http.client.HTTPConnection('127.0.0.1', port, timeout=10)
        conn.request(
            'POST',
            '/messages/?session_id=nonexistent',
            body=json.dumps(INITIALIZE),
            headers={'Content-Type': 'application/json'},
        )
        resp = conn.getresponse()
        assert resp.status == 400
        conn.close()
    finally:
        _stop_server(proc)


def test_cors_preflight():
    """OPTIONS preflight returns CORS headers."""
    port = 18767
    proc = _start_server(port)
    try:
        conn = http.client.HTTPConnection('127.0.0.1', port, timeout=10)
        conn.request(
            'OPTIONS',
            '/sse',
            headers={
                'Origin': 'http://localhost:3000',
                'Access-Control-Request-Method': 'GET',
            },
        )
        resp = conn.getresponse()
        assert resp.status == 200
        assert resp.getheader('Access-Control-Allow-Origin') == '*'
        assert resp.getheader('Access-Control-Allow-Methods')
        conn.close()
    finally:
        _stop_server(proc)


if __name__ == '__main__':
    tests = [
        test_sse_server_lifecycle,
        test_missing_session_rejected,
        test_cors_preflight,
    ]
    passed = 0
    for t in tests:
        try:
            t()
            passed += 1
            print(f'  PASS  {t.__name__}')
        except Exception as e:
            print(f'  FAIL  {t.__name__}: {e}')
    print(f'\n{passed}/{len(tests)} passed')
