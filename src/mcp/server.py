"""TianshangScribe MCP Server — stdio / SSE / Streamable HTTP.

Built on the official MCP Python SDK (:class:`mcp.server.mcpserver.MCPServer`).
Each of the seven document tools is registered from its function signature
(the SDK derives the ``inputSchema`` from the ``Annotated`` parameters), and
the transports are wired in :mod:`src.mcp.transport`.

Usage:
  python -m src.mcp.server                            # stdio (default)
  python -m src.mcp.server --transport sse --port 8080
  python -m src.mcp.server --transport streamable-http --port 8080
"""

from __future__ import annotations

import argparse

from mcp.server.mcpserver import MCPServer
from mcp_types import ToolAnnotations

from src.mcp import transport
from src.mcp.metrics import instrumented
from src.mcp.prompts import register_prompts
from src.mcp.rate_limit import RateLimiter
from src.mcp.security import is_destructive, is_idempotent, is_read_only
from src.mcp.tool_search import install_tool_search
from src.mcp.tools._registry import get_tools
from src.utils.config import Settings
from src.utils.logging import configure_logging

SERVER_NAME = 'tianshang-scribe'
SERVER_VERSION = '0.3.0'

INSTRUCTIONS = (
    'TianshangScribe MCP Server — create, edit, fill, convert and extract '
    'Office documents (Word, Excel, PowerPoint). LaTeX-style markup and '
    'math formulas are supported.\n'
    'Tools that accept output_path write files to disk; edit_office_document '
    'overwrites its input when output_path is omitted. extract_document_data '
    'and validate_template are read-only and never modify their input files; '
    'compare_documents is read-only in compare mode but writes to the snapshot '
    'store in snapshot/restore mode.\n'
    'HTTP transports may require an Authorization: Bearer <key> header '
    '(configured via SCRIBE_API_KEYS or --auth-token) and are subject to '
    'per-client rate limits.'
)


def build_server(version: str = SERVER_VERSION) -> MCPServer:
    """Build and configure the MCP server with tools, prompts and routes."""
    server = MCPServer(name=SERVER_NAME, version=version, instructions=INSTRUCTIONS)

    for entry in get_tools():
        name = entry['name']
        annotations = ToolAnnotations(
            read_only_hint=is_read_only(name),
            destructive_hint=is_destructive(name),
            idempotent_hint=is_idempotent(name),
            open_world_hint=False,
        )
        server.add_tool(
            instrumented(entry['fn'], name),
            name=name,
            description=entry['description'],
            annotations=annotations,
        )

    register_prompts(server)
    install_tool_search(server)
    transport.register_observability(server, version)
    return server


def main() -> None:
    """CLI entry point (``scribe-mcp`` / ``python -m src.mcp.server``)."""
    settings = Settings()
    configure_logging(settings)
    parser = argparse.ArgumentParser(description='TianshangScribe MCP Server')
    parser.add_argument(
        '--transport',
        choices=['stdio', 'sse', 'streamable-http'],
        default=settings.transport,
        help='Transport protocol (default: %(default)s)',
    )
    parser.add_argument(
        '--host',
        default=settings.host,
        help='HTTP host (default: %(default)s)',
    )
    parser.add_argument(
        '--port',
        type=int,
        default=settings.port,
        help='HTTP port (default: %(default)s)',
    )
    parser.add_argument(
        '--auth-token',
        default=settings.auth_token,
        help='Bearer token for HTTP auth (env: SCRIBE_AUTH_TOKEN)',
    )
    parser.add_argument(
        '--cors-origins',
        default=settings.cors_origins,
        help='CORS allowed origins, comma-separated (env: SCRIBE_CORS_ORIGINS)',
    )
    parser.add_argument(
        '--rate-limit-max',
        type=int,
        default=settings.rate_limit_max,
        help='Max requests per client per window (default: %(default)s)',
    )
    parser.add_argument(
        '--rate-limit-window',
        type=int,
        default=settings.rate_limit_window,
        help='Rate limit window in seconds (default: %(default)s)',
    )
    parser.add_argument(
        '--mcp-path',
        default=settings.mcp_path,
        help='Streamable HTTP endpoint path (default: %(default)s)',
    )
    args = parser.parse_args()

    server = build_server()
    auth_tokens = ','.join(filter(None, [args.auth_token, settings.api_keys]))
    limiter = RateLimiter(
        max_requests=args.rate_limit_max,
        window_seconds=args.rate_limit_window,
    )

    if args.transport == 'stdio':
        transport.run_stdio(server)
        return

    if args.transport == 'sse':
        transport.run_sse(
            server,
            host=args.host,
            port=args.port,
            auth_token=auth_tokens,
            cors_origins=args.cors_origins,
            rate_limiter=limiter,
        )
        return

    transport.run_http(
        server,
        host=args.host,
        port=args.port,
        auth_token=auth_tokens,
        cors_origins=args.cors_origins,
        rate_limiter=limiter,
        streamable_http_path=args.mcp_path,
    )


if __name__ == '__main__':
    main()
