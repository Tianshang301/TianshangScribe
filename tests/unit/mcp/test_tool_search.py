"""Unit tests for SEP-1821 Tool Search (``tools/list`` ``query`` parameter)."""

from __future__ import annotations

import asyncio

import pytest
from mcp_types import Tool

from tianshang_scribe.mcp.server import build_server
from tianshang_scribe.mcp.tool_search import (
    TOOLS_LIST_METHOD,
    ToolSearchParams,
    install_tool_search,
    score_tool,
    search_tools,
)


def _tool(name: str, description: str = '', title: str | None = None) -> Tool:
    return Tool(
        name=name,
        title=title,
        description=description or f'Does {name} things.',
        input_schema={'type': 'object', 'properties': {'a': {'type': 'string'}}},
    )


TOOLS = [
    _tool('create_office_document'),
    _tool('edit_office_document'),
    _tool('convert_document'),
    _tool('extract_text', title='Text Extraction'),
    _tool('validate_template'),
]


def tool_names(result: list[Tool]) -> list[str]:
    return [t.name for t in result]


class TestScoreTool:
    def test_empty_query_scores_zero(self) -> None:
        assert score_tool(TOOLS[0], '') == 0
        assert score_tool(TOOLS[0], '   ') == 0

    def test_name_match_is_weighted_higher_than_description_only(self) -> None:
        # 'create' in the name scores 8; the same word only in the
        # description of another tool scores 2.
        name_match = _tool('create', description='unrelated.')
        desc_only = _tool('doc', description='create convert text thing')
        assert score_tool(name_match, 'create') == 8
        assert score_tool(desc_only, 'create') == 2

    def test_description_match_scores_positive(self) -> None:
        desc_only = _tool('doc', description='create convert text')
        assert score_tool(desc_only, 'convert') == 2
        assert score_tool(desc_only, 'convert') > 0

    def test_no_match_scores_zero(self) -> None:
        assert score_tool(TOOLS[0], 'zzzznotpresent') == 0

    def test_title_matches(self) -> None:
        assert score_tool(TOOLS[3], 'text') > 0

    def test_case_insensitive(self) -> None:
        assert score_tool(TOOLS[0], 'CREATE') == score_tool(TOOLS[0], 'create')


class TestSearchTools:
    def test_none_or_blank_query_returns_all(self) -> None:
        for query in [None, '', '  ']:
            assert tool_names(search_tools(TOOLS, query)) == tool_names(TOOLS)

    def test_substring_name_filter(self) -> None:
        result = tool_names(search_tools(TOOLS, 'office'))
        assert result[0] == 'create_office_document'
        assert 'edit_office_document' in result

    def test_ranks_name_before_description(self) -> None:
        result = search_tools(TOOLS, 'extract')
        assert result[0].name == 'extract_text'
        assert 'extract_text' in tool_names(result)

    def test_no_matches_returns_empty(self) -> None:
        assert search_tools(TOOLS, 'zzzznope') == []

    def test_stable_tie_ordering(self) -> None:
        # Both document tools match 'document' in the name; registration order holds
        result = tool_names(search_tools(TOOLS, 'document'))
        assert result == ['create_office_document', 'edit_office_document', 'convert_document']

    def test_partial_token_matching(self) -> None:
        assert 'validate_template' in tool_names(search_tools(TOOLS, 'valid'))


class TestRegisteredHandler:
    def test_handler_is_replaced_with_search_params(self) -> None:
        server = build_server()
        entry = server._lowlevel_server.get_request_handler(TOOLS_LIST_METHOD)
        assert entry is not None
        assert entry.params_type is ToolSearchParams

    def test_no_query_returns_full_list(self) -> None:
        server = build_server()
        handler = server._lowlevel_server.get_request_handler(TOOLS_LIST_METHOD).handler

        async def run() -> list[str]:
            result = await handler(None, ToolSearchParams())
            return [t.name for t in result.tools]

        names = asyncio.run(run())
        assert names == [
            'create_office_document',
            'edit_office_document',
            'fill_template',
            'convert_document',
            'extract_document_data',
            'validate_template',
            'compare_documents',
            'create_excel_workbook',
            'edit_excel_workbook',
            'create_presentation',
            'edit_presentation',
            'analyze_excel_data',
        ]

    def test_query_returns_matching_subset(self) -> None:
        server = build_server()
        handler = server._lowlevel_server.get_request_handler(TOOLS_LIST_METHOD).handler

        async def run() -> list[str]:
            result = await handler(None, ToolSearchParams(query='pdf'))
            return [t.name for t in result.tools]

        assert asyncio.run(run()) == ['convert_document']

    def test_query_none_params_defaults_to_full(self) -> None:
        server = build_server()
        handler = server._lowlevel_server.get_request_handler(TOOLS_LIST_METHOD).handler

        async def run() -> int:
            result = await handler(None, None)
            return len(result.tools)

        assert asyncio.run(run()) == 12

    def test_install_is_idempotent(self) -> None:
        server = build_server()
        install_tool_search(server)
        install_tool_search(server)
        entry = server._lowlevel_server.get_request_handler(TOOLS_LIST_METHOD)
        assert entry is not None
        assert entry.params_type is ToolSearchParams

    def test_serialized_tools_are_valid_tool_models(self) -> None:
        server = build_server()
        handler = server._lowlevel_server.get_request_handler(TOOLS_LIST_METHOD).handler

        async def run() -> Tool:
            result = await handler(None, ToolSearchParams(query='extract'))
            return result.tools[0]

        tool = asyncio.run(run())
        assert isinstance(tool, Tool)
        assert tool.name == 'extract_document_data'
        assert tool.description
        assert tool.input_schema


class TestParamsValidation:
    def test_model_rejects_unknown_required(self) -> None:
        with pytest.raises(ValueError):
            ToolSearchParams(cursor=123)  # type: ignore[arg-type]

    def test_query_is_optional(self) -> None:
        params = ToolSearchParams()
        assert params.query is None
        assert params.cursor is None
