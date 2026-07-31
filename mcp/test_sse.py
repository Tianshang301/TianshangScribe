"""Tests for SSE MCP Server transport."""
import http.client
import json
import subprocess
import sys
import time
from pathlib import Path

PROJECT_ROOT = str(Path(__file__).resolve().parent.parent)


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
        [sys.executable, '-m', 'mcp.server', '--transport', 'sse', '--port', str(port)],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        cwd=PROJECT_ROOT,
    )
    _wait_for_server('127.0.0.1', port)
    return proc


def test_sse_server_lifecycle():
    """SSE server handles initialize + tools/list + tools/call."""
    port = 18765
    proc = _start_server(port)
    sse_conn = None
    try:
        sse_conn = http.client.HTTPConnection('127.0.0.1', port, timeout=15)
        sse_conn.request('GET', '/sse')
        sse_resp = sse_conn.getresponse()
        assert sse_resp.status == 200

        sse_data = b''
        while b'\n\n' not in sse_data:
            chunk = sse_resp.read(1)
            if not chunk:
                break
            sse_data += chunk

        sse_text = sse_data.decode()
        assert 'event: endpoint' in sse_text

        session_id = ''
        for line in sse_text.split('\n'):
            if line.startswith('data: ') and 'session_id=' in line:
                session_id = line.split('session_id=')[1].rstrip()
                break
        assert session_id, 'Could not extract session_id from endpoint event'

        # -- initialize --
        conn = http.client.HTTPConnection('127.0.0.1', port, timeout=10)
        req = json.dumps({
            'jsonrpc': '2.0', 'id': 1, 'method': 'initialize',
            'params': {'protocolVersion': '2024-11-05', 'capabilities': {}},
        })
        conn.request(
            'POST', f'/message?session_id={session_id}', body=req,
            headers={'Content-Type': 'application/json'},
        )
        resp = conn.getresponse()
        body = json.loads(resp.read().decode())
        assert body['result']['serverInfo']['name'] == 'tianshang-scribe'
        conn.close()

        # -- tools/list --
        conn = http.client.HTTPConnection('127.0.0.1', port, timeout=10)
        req = json.dumps({
            'jsonrpc': '2.0', 'id': 2, 'method': 'tools/list', 'params': {},
        })
        conn.request(
            'POST', f'/message?session_id={session_id}', body=req,
            headers={'Content-Type': 'application/json'},
        )
        resp = conn.getresponse()
        body = json.loads(resp.read().decode())
        assert len(body['result']['tools']) == 7
        conn.close()

        # -- tools/call create_office_document --
        conn = http.client.HTTPConnection('127.0.0.1', port, timeout=10)
        req = json.dumps({
            'jsonrpc': '2.0', 'id': 3, 'method': 'tools/call',
            'params': {
                'name': 'create_office_document',
                'arguments': {
                    'format': 'docx',
                    'content': [
                        {'type': 'heading', 'text': 'Test', 'level': 1},
                    ],
                },
            },
        })
        conn.request(
            'POST', f'/message?session_id={session_id}', body=req,
            headers={'Content-Type': 'application/json'},
        )
        resp = conn.getresponse()
        body = json.loads(resp.read().decode())
        content = body['result']['content'][0]['text']
        result = json.loads(content)
        assert result['success']
        conn.close()

    finally:
        if sse_conn:
            sse_conn.close()
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            proc.kill()
            proc.wait()


def test_missing_session_rejected():
    """POST /message with invalid session_id returns 400."""
    port = 18766
    proc = _start_server(port)
    try:
        conn = http.client.HTTPConnection('127.0.0.1', port, timeout=10)
        req = json.dumps({
            'jsonrpc': '2.0', 'id': 1, 'method': 'initialize', 'params': {},
        })
        conn.request(
            'POST', '/message?session_id=nonexistent', body=req,
            headers={'Content-Type': 'application/json'},
        )
        resp = conn.getresponse()
        assert resp.status == 400
        conn.close()
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            proc.kill()
            proc.wait()


def test_cors_preflight():
    """OPTIONS request returns 204 with CORS headers."""
    port = 18767
    proc = _start_server(port)
    try:
        conn = http.client.HTTPConnection('127.0.0.1', port, timeout=10)
        conn.request('OPTIONS', '/sse')
        resp = conn.getresponse()
        assert resp.status == 204
        assert resp.getheader('Access-Control-Allow-Origin') == '*'
        conn.close()
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            proc.kill()
            proc.wait()


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
