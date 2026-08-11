"""MCP tool permission classification.

Each tool is assigned a permission level. Read-only tools (extract / validate /
compare) are auto-approved; write tools (create / fill / convert) require
confirmation; destructive tools (edit, which may overwrite its input) must
always be confirmed. The levels drive the SDK ``ToolAnnotations`` exposed in
``tools/list`` as well as any future approval gate.
"""

from __future__ import annotations

from enum import Enum


class PermissionLevel(Enum):
    """Permission levels for MCP tools."""

    READ_ONLY = 'read_only'
    STANDARD = 'standard'
    DESTRUCTIVE = 'destructive'
    ADMIN = 'admin'


#: Permission level for every registered tool.
TOOL_PERMISSIONS: dict[str, PermissionLevel] = {
    'create_office_document': PermissionLevel.STANDARD,
    'edit_office_document': PermissionLevel.DESTRUCTIVE,
    'fill_template': PermissionLevel.STANDARD,
    'convert_document': PermissionLevel.STANDARD,
    'extract_document_data': PermissionLevel.READ_ONLY,
    'validate_template': PermissionLevel.READ_ONLY,
    'compare_documents': PermissionLevel.READ_ONLY,
}


def levels_for(tool_name: str) -> set[PermissionLevel]:
    """Return the set of permission levels the tool may exercise.

    Unknown tools default to ``STANDARD``.
    """
    return {TOOL_PERMISSIONS.get(tool_name, PermissionLevel.STANDARD)}


def is_read_only(tool_name: str) -> bool:
    """Return whether every sub-operation of the tool is read-only."""
    return levels_for(tool_name) == {PermissionLevel.READ_ONLY}


def is_destructive(tool_name: str) -> bool:
    """Return whether the tool may destroy data."""
    return PermissionLevel.DESTRUCTIVE in levels_for(tool_name)


def is_idempotent(tool_name: str) -> bool:
    """Return whether repeated identical calls produce identical results.

    Read-only tools are idempotent; file-writing tools are not.
    """
    return is_read_only(tool_name)


def check_permission(tool_name: str, auto_approve: set[str]) -> bool:
    """Return whether ``tool_name`` may auto-execute.

    Tools are allowed when they are listed in ``auto_approve`` or when every
    sub-operation is read-only. Unknown tools default to ``STANDARD``.
    """
    if tool_name in auto_approve:
        return True
    return is_read_only(tool_name)
