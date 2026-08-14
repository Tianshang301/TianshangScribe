"""SEP-1821 Tool Search — optional ``query`` filtering for ``tools/list``.

The MCP spec lets a client pass a ``query`` member to ``tools/list`` to
request a relevance-filtered view of the tool catalogue instead of the full
list. The SDK v2 handler binds ``tools/list`` to ``PaginatedRequestParams``
and ignores the extra member, so this module registers a replacement handler
that accepts :class:`ToolSearchParams` (a ``PaginatedRequestParams`` subclass
gaining ``query``) and returns only tools whose name, description or input
schema match the query, ranked by relevance.

Registration replaces the stock handler via the SDK's documented
``Server.add_request_handler`` replace semantics (the same mechanism the
``Extension.methods()`` API uses for additive methods).
"""

from __future__ import annotations

from collections.abc import Sequence
from typing import Any

from mcp.server.context import ServerRequestContext
from mcp.server.lowlevel.server import Server as LowLevelServer
from mcp.server.mcpserver import MCPServer
from mcp_types import ListToolsResult, PaginatedRequestParams, Tool

TOOLS_LIST_METHOD = 'tools/list'


class ToolSearchParams(PaginatedRequestParams):
    """Parameters for ``tools/list`` with the SEP-1821 ``query`` extension.

    Extends the SDK's pagination model so the ``query`` member survives
    validation and reaches the handler. ``cursor`` stays as an optional
    opaque pagination token (unused by this server, kept for API parity).
    """

    query: str | None = None
    """Free-text search query. When omitted, the full tool list is returned."""


def _terms(query: str) -> list[str]:
    """Split ``query`` into lowercase token terms."""
    return [token for token in query.lower().replace('-', ' _ ').replace('_', ' ').split() if token]


def _schema_text(schema: dict[str, Any]) -> str:
    """Flatten an input schema into lowercase searchable text."""
    text: list[str] = []
    if isinstance(schema.get('title'), str):
        text.append(schema['title'])
    properties = schema.get('properties')
    if isinstance(properties, dict):
        for name, spec in properties.items():
            text.append(name)
            if isinstance(spec, dict):
                if isinstance(spec.get('title'), str):
                    text.append(spec['title'])
                if isinstance(spec.get('description'), str):
                    text.append(spec['description'])
    return ' '.join(text).lower()


def score_tool(tool: Tool, query: str) -> int:
    """Return a relevance score for one tool against ``query``.

    Weights descend from name (strongest signal) through description to the
    input property names. A token must be present in at least name,
    description or schema text to contribute; a non-matching query scores 0.
    """
    terms = _terms(query)
    if not terms:
        return 0

    name = (tool.name or '').lower()
    title = (tool.title or '').lower()
    description = (tool.description or '').lower()
    schema = _schema_text(tool.input_schema)

    score = 0
    for term in terms:
        if term in name:
            score += 8
        elif term in title:
            score += 4
        if term in description:
            score += 2
        elif term in schema:
            score += 1
    return score


def search_tools(tools: Sequence[Tool], query: str | None) -> list[Tool]:
    """Filter and rank ``tools`` by ``query`` relevance.

    With no/blank ``query`` the tools are returned unchanged in registration
    order. Otherwise only tools with a positive score are kept, ordered by
    descending score, ties broken by original registration order (stable).
    """
    if not query or not query.strip():
        return list(tools)

    scored = [(score_tool(tool, query), index, tool) for index, tool in enumerate(tools)]
    matched = [(score, index, tool) for score, index, tool in scored if score > 0]
    matched.sort(key=lambda item: (-item[0], item[1]))
    return [tool for _, _, tool in matched]


async def _handle_list_tools_search(
    ctx: ServerRequestContext[Any],
    params: ToolSearchParams | None,
    server: MCPServer,
) -> ListToolsResult:
    """Serve ``tools/list`` honouring the optional ``query`` parameter."""
    tools = await server.list_tools()
    return ListToolsResult(tools=search_tools(tools, params.query if params else None))


def install_tool_search(server: MCPServer) -> None:
    """Register the ``tools/list`` handler with SEP-1821 ``query`` support.

    Replaces the SDK's stock handler on the wrapped low-level server. The
    response for clients that never send ``query`` is byte-identical to the
    stock behaviour (full list, registration order).
    """
    lowlevel: LowLevelServer[Any] = server._lowlevel_server

    async def handler(
        ctx: ServerRequestContext[Any], params: ToolSearchParams | None
    ) -> ListToolsResult:
        return await _handle_list_tools_search(ctx, params, server)

    lowlevel.add_request_handler(TOOLS_LIST_METHOD, ToolSearchParams, handler)
