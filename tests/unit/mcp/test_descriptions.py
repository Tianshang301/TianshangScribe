"""MCP tool description / annotation consistency tests.

Guards the principle that a tool's ``description`` must disclose its side
effects consistently with the ``readOnlyHint`` / ``destructiveHint``
annotations derived in :mod:`src.mcp.security`.
"""

from __future__ import annotations

from src.mcp.security import is_destructive, is_read_only
from src.mcp.tools._registry import TOOLS

WRITE_TOOLS = {
    'create_office_document',
    'edit_office_document',
    'fill_template',
    'convert_document',
}
READ_ONLY_TOOLS = {'extract_document_data', 'validate_template', 'compare_documents'}
DESTRUCTIVE_TOOLS = {'edit_office_document'}


def _desc(name: str) -> str:
    return next(t['description'] for t in TOOLS if t['name'] == name)


def test_registry_annotations_consistent() -> None:
    names = {t['name'] for t in TOOLS}
    assert names == (WRITE_TOOLS | READ_ONLY_TOOLS)
    for name in names:
        assert is_read_only(name) == (name in READ_ONLY_TOOLS), name
        assert is_destructive(name) == (name in DESTRUCTIVE_TOOLS), name


def test_write_tools_disclose_side_effects() -> None:
    for name in WRITE_TOOLS:
        desc = _desc(name).lower()
        assert 'writes' in desc or 'overwrite' in desc, name


def test_read_only_tools_affirm_no_modification() -> None:
    for name in READ_ONLY_TOOLS:
        desc = _desc(name).lower()
        assert 'read-only' in desc and 'never modifies' in desc, name


def test_edit_discloses_in_place_overwrite() -> None:
    assert 'overwritten in place' in _desc('edit_office_document').lower()
    assert 'backup' in _desc('edit_office_document').lower()


def test_convert_discloses_pdf_engine_dependency() -> None:
    desc = _desc('convert_document').lower()
    assert 'office2pdf' in desc and 'libreoffice' in desc


def test_validate_guides_call_order() -> None:
    desc = _desc('validate_template').lower()
    assert 'before' in desc and 'fill_template' in desc


def test_descriptions_mention_sibling_tool() -> None:
    """Each description points at a sibling tool (dimension: when to use what)."""
    all_names = WRITE_TOOLS | READ_ONLY_TOOLS
    for tool in TOOLS:
        siblings = all_names - {tool['name']}
        desc = tool['description']
        assert any(s in desc for s in siblings), tool['name']


def test_descriptions_tight_length() -> None:
    """Three-sentence template: ≤ 90 words, at least two sentences."""
    for tool in TOOLS:
        words = len(tool['description'].split())
        assert words <= 90, (tool['name'], words)
        assert tool['description'].count('.') >= 2, tool['name']
