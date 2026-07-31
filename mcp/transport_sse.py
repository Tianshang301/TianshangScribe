"""SSE transport for TianshangScribe MCP Server.

Provides HTTP+Server-Sent Events transport for cloud Agent platforms.
Supports Bearer Token auth, CORS whitelist, health checks, and structured logging.
Zero external dependencies — pure asyncio + stdlib.
"""

from __future__ import annotations

import asyncio
import json
import logging
import sys
import time
import uuid
from http import HTTPStatus
from typing import Optional

from mcp.server import SERVER_NAME, SERVER_VERSION, TOOLS, _handle_request

_START_TIME = time.time()
_logger = logging.getLogger('tianshang-scribe')


def _setup_structured_logging() -> None:
    handler = logging.StreamHandler(sys.stderr)
    handler.setFormatter(
        logging.Formatter(
            '{"timestamp":"%(asctime)s","level":"%(levelname)s",'
            '"logger":"%(name)s","message":%(message)s}',
            datefmt='%Y-%m-%dT%H:%M:%S',
        )
    )
    _logger.addHandler(handler)
    _logger.setLevel(logging.INFO)
    _logger.propagate = False


def _log_event(event: str, **kwargs) -> None:
    payload = json.dumps({**kwargs, 'event': event}, ensure_ascii=False)
    _logger.info(payload)


class SseServer:
    """Async HTTP server with SSE transport for MCP JSON-RPC."""

    def __init__(
        self,
        host: str = '127.0.0.1',
        port: int = 8080,
        auth_token: Optional[str] = None,
        cors_origins: Optional[str] = None,
    ):
        self.host = host
        self.port = port
        self._auth_token = auth_token
        self._cors_origins: list[str] = []
        if cors_origins:
            self._cors_origins = [
                o.strip() for o in cors_origins.split(',') if o.strip()
            ]
        self._sessions: dict[str, asyncio.Queue] = {}

    @property
    def _cors_header(self) -> str:
        if not self._cors_origins:
            return '*'
        return ', '.join(self._cors_origins)

    @property
    def _active_sessions(self) -> int:
        return len(self._sessions)

    async def start(self) -> None:
        """Start the SSE HTTP server."""
        server = await asyncio.start_server(
            self._handle_client, self.host, self.port,
        )
        _log_event('server_start', host=self.host, port=self.port,
                   auth=(self._auth_token is not None),
                   cors=self._cors_header)
        print(f'SSE MCP Server listening on http://{self.host}:{self.port}/sse')
        if self._auth_token:
            print('  Auth: Bearer Token enabled')
        if self._cors_origins:
            print(f'  CORS: {self._cors_header}')
        async with server:
            await server.serve_forever()

    async def _handle_client(
        self, reader: asyncio.StreamReader, writer: asyncio.StreamWriter,
    ) -> None:
        t0 = time.time()
        method = ''
        path_only = ''
        try:
            request_line = await asyncio.wait_for(reader.readline(), timeout=30)
            if not request_line:
                return

            parts = request_line.decode().strip().split()
            if len(parts) < 2:
                self._write_response(writer, 400, '{"error":"Bad Request"}')
                await writer.drain()
                return
            method, path = parts[0], parts[1]

            headers: dict[str, str] = {}
            while True:
                header_line = await reader.readline()
                if not header_line or header_line in (b'\r\n', b'\n'):
                    break
                key, sep, value = header_line.decode().strip().partition(':')
                if sep:
                    headers[key.strip().lower()] = value.strip()

            body = ''
            if method == 'POST':
                content_length = int(headers.get('content-length', 0))
                if content_length > 0:
                    body = (await reader.readexactly(content_length)).decode()

            path_only = path.split('?')[0]
            query: dict[str, str] = {}
            if '?' in path:
                for pair in path.split('?', 1)[1].split('&'):
                    if '=' in pair:
                        k, v = pair.split('=', 1)
                        query[k] = v

            if method == 'OPTIONS':
                self._write_cors_preflight(writer)
                await writer.drain()
            elif path_only == '/health' and method == 'GET':
                self._write_health(writer)
                await writer.drain()
            elif path_only == '/sse' and method == 'GET':
                await self._handle_sse(writer)
            elif path_only == '/message' and method == 'POST':
                if self._auth_token:
                    auth_header = headers.get('authorization', '')
                    if not auth_header.startswith('Bearer '):
                        self._write_response(
                            writer, 401,
                            '{"error":"Unauthorized: Bearer token required"}')
                        await writer.drain()
                        _log_event('auth_failed', reason='missing_token',
                                   path=path_only, method=method)
                        return
                    if auth_header[7:] != self._auth_token:
                        self._write_response(
                            writer, 403,
                            '{"error":"Forbidden: invalid token"}')
                        await writer.drain()
                        _log_event('auth_failed', reason='invalid_token',
                                   path=path_only, method=method)
                        return
                session_id = query.get('session_id', '')
                await self._handle_message(writer, body, session_id)
            else:
                self._write_response(writer, 404, '{"error":"Not Found"}')
                await writer.drain()

        except (ConnectionResetError, BrokenPipeError, asyncio.TimeoutError):
            pass
        except Exception:
            try:
                self._write_response(writer, 500,
                                     '{"error":"Internal Server Error"}')
                await writer.drain()
            except Exception:
                pass
        finally:
            dt_ms = int((time.time() - t0) * 1000)
            if method and path_only:
                _log_event('request', method=method, path=path_only,
                           duration_ms=dt_ms)
            try:
                writer.close()
                await writer.wait_closed()
            except Exception:
                pass

    async def _handle_sse(self, writer: asyncio.StreamWriter) -> None:
        """Handle GET /sse — establish SSE connection."""
        session_id = str(uuid.uuid4())
        queue: asyncio.Queue = asyncio.Queue()
        self._sessions[session_id] = queue

        _log_event('sse_connect', session_id=session_id,
                   total_sessions=self._active_sessions)

        writer.write(b'HTTP/1.1 200 OK\r\n')
        writer.write(b'Content-Type: text/event-stream\r\n')
        writer.write(b'Cache-Control: no-cache\r\n')
        writer.write(b'Connection: keep-alive\r\n')
        writer.write(
            f'Access-Control-Allow-Origin: {self._cors_header}\r\n'.encode()
        )
        writer.write(b'\r\n')
        await writer.drain()

        await self._send_sse(writer, 'endpoint',
                             f'/message?session_id={session_id}')

        try:
            while True:
                try:
                    event_data = await asyncio.wait_for(queue.get(), timeout=30)
                except asyncio.TimeoutError:
                    await self._send_sse(writer, 'heartbeat', '{}')
                    continue
                if event_data is None:
                    break
                await self._send_sse(writer, 'message', event_data)
        except (ConnectionResetError, BrokenPipeError):
            pass
        finally:
            self._sessions.pop(session_id, None)
            _log_event('sse_disconnect', session_id=session_id,
                       total_sessions=self._active_sessions)

    async def _handle_message(
        self, writer: asyncio.StreamWriter, body: str, session_id: str,
    ) -> None:
        """Handle POST /message — JSON-RPC request -> response."""
        if not session_id or session_id not in self._sessions:
            self._write_response(writer, 400,
                                 '{"error":"Missing or invalid session_id"}')
            await writer.drain()
            return

        try:
            request = json.loads(body)
        except json.JSONDecodeError:
            self._write_response(writer, 400, '{"error":"Invalid JSON"}')
            await writer.drain()
            return

        req_method = request.get('method', '')
        req_id = request.get('id')
        _log_event('rpc_call', session_id=session_id[:8],
                   method=req_method, id=req_id)

        response = _handle_request(request)

        if response is not None:
            response_text = json.dumps(response, ensure_ascii=False)
            await self._sessions[session_id].put(response_text)
            self._write_response(writer, 200, response_text)
        else:
            self._write_response(writer, 202, '{}')
        await writer.drain()

    async def _send_sse(
        self, writer: asyncio.StreamWriter, event: str, data: str,
    ) -> None:
        """Send an SSE-formatted event."""
        message = f'event: {event}\ndata: {data}\n\n'
        writer.write(message.encode())
        await writer.drain()

    def _write_response(
        self,
        writer: asyncio.StreamWriter,
        status: int,
        body: str,
        content_type: str = 'application/json',
    ) -> None:
        """Write HTTP response (caller must drain)."""
        body_bytes = body.encode()
        writer.write(
            f'HTTP/1.1 {status} {HTTPStatus(status).phrase}\r\n'.encode()
        )
        writer.write(f'Content-Type: {content_type}\r\n'.encode())
        writer.write(f'Content-Length: {len(body_bytes)}\r\n'.encode())
        writer.write(
            f'Access-Control-Allow-Origin: {self._cors_header}\r\n'.encode()
        )
        writer.write(b'\r\n')
        writer.write(body_bytes)

    def _write_cors_preflight(self, writer: asyncio.StreamWriter) -> None:
        """Respond to CORS preflight (OPTIONS)."""
        writer.write(b'HTTP/1.1 204 No Content\r\n')
        writer.write(
            f'Access-Control-Allow-Origin: {self._cors_header}\r\n'.encode()
        )
        writer.write(b'Access-Control-Allow-Methods: GET, POST, OPTIONS\r\n')
        writer.write(
            b'Access-Control-Allow-Headers: '
            b'Content-Type, Authorization\r\n'
        )
        writer.write(b'Content-Length: 0\r\n')
        writer.write(b'\r\n')

    def _write_health(self, writer: asyncio.StreamWriter) -> None:
        """Handle GET /health — return server status."""
        from src.transform.pdf import _find_office2pdf
        pdf_engine = 'office2pdf' if _find_office2pdf() else 'none'
        payload = json.dumps({
            'status': 'ok',
            'version': SERVER_VERSION,
            'name': SERVER_NAME,
            'uptime_seconds': int(time.time() - _START_TIME),
            'active_sessions': self._active_sessions,
            'tools_available': len(TOOLS),
            'tools': [t['name'] for t in TOOLS],
            'pdf_engine': pdf_engine,
            'auth_enabled': self._auth_token is not None,
        }, ensure_ascii=False)
        self._write_response(writer, 200, payload)


def run_sse(
    host: str = '127.0.0.1',
    port: int = 8080,
    auth_token: Optional[str] = None,
    cors_origins: Optional[str] = None,
) -> None:
    """Entry point: start SSE server."""
    _setup_structured_logging()
    server = SseServer(host=host, port=port, auth_token=auth_token,
                       cors_origins=cors_origins)
    asyncio.run(server.start())


if __name__ == '__main__':
    run_sse()
