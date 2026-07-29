"""SSE transport for TianshangScribe MCP Server.

Provides HTTP+Server-Sent Events transport for cloud Agent platforms.
Zero external dependencies — pure asyncio + stdlib.
"""

from __future__ import annotations

import asyncio
import json
import uuid
from http import HTTPStatus

from mcp.server import _handle_request


class SseServer:
    """Async HTTP server with SSE transport for MCP JSON-RPC."""

    def __init__(self, host: str = '127.0.0.1', port: int = 8080):
        self.host = host
        self.port = port
        self._sessions: dict[str, asyncio.Queue] = {}

    async def start(self) -> None:
        """Start the SSE HTTP server."""
        server = await asyncio.start_server(
            self._handle_client, self.host, self.port,
        )
        print(f'SSE MCP Server listening on http://{self.host}:{self.port}/sse')
        async with server:
            await server.serve_forever()

    async def _handle_client(
        self, reader: asyncio.StreamReader, writer: asyncio.StreamWriter,
    ) -> None:
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
            elif path_only == '/sse' and method == 'GET':
                await self._handle_sse(writer)
            elif path_only == '/message' and method == 'POST':
                session_id = query.get('session_id', '')
                await self._handle_message(writer, body, session_id)
            else:
                self._write_response(writer, 404, '{"error":"Not Found"}')
                await writer.drain()
        except (ConnectionResetError, BrokenPipeError, asyncio.TimeoutError):
            pass
        except Exception:
            try:
                self._write_response(writer, 500, '{"error":"Internal Server Error"}')
                await writer.drain()
            except Exception:
                pass
        finally:
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

        writer.write(b'HTTP/1.1 200 OK\r\n')
        writer.write(b'Content-Type: text/event-stream\r\n')
        writer.write(b'Cache-Control: no-cache\r\n')
        writer.write(b'Connection: keep-alive\r\n')
        writer.write(b'Access-Control-Allow-Origin: *\r\n')
        writer.write(b'\r\n')
        await writer.drain()

        await self._send_sse(writer, 'endpoint', f'/message?session_id={session_id}')

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

    async def _handle_message(
        self, writer: asyncio.StreamWriter, body: str, session_id: str,
    ) -> None:
        """Handle POST /message — JSON-RPC request -> response."""
        if not session_id or session_id not in self._sessions:
            self._write_response(writer, 400, '{"error":"Missing or invalid session_id"}')
            await writer.drain()
            return

        try:
            request = json.loads(body)
        except json.JSONDecodeError:
            self._write_response(writer, 400, '{"error":"Invalid JSON"}')
            await writer.drain()
            return

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
        writer.write(b'Access-Control-Allow-Origin: *\r\n')
        writer.write(b'\r\n')
        writer.write(body_bytes)

    def _write_cors_preflight(self, writer: asyncio.StreamWriter) -> None:
        """Respond to CORS preflight (OPTIONS)."""
        writer.write(b'HTTP/1.1 204 No Content\r\n')
        writer.write(b'Access-Control-Allow-Origin: *\r\n')
        writer.write(b'Access-Control-Allow-Methods: GET, POST, OPTIONS\r\n')
        writer.write(b'Access-Control-Allow-Headers: Content-Type\r\n')
        writer.write(b'Content-Length: 0\r\n')
        writer.write(b'\r\n')


def run_sse(host: str = '127.0.0.1', port: int = 8080) -> None:
    """Entry point: start SSE server."""
    server = SseServer(host, port)
    asyncio.run(server.start())


if __name__ == '__main__':
    run_sse()
