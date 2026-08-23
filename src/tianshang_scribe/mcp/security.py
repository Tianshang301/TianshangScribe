"""MCP tool permission classification and role-based access control.

Each tool is assigned a permission level. Read-only tools (extract / validate)
are auto-approved; standard tools (create / fill / convert / compare, whose
snapshot sub-operations write) require confirmation; destructive tools (edit,
which may overwrite its input) must always be confirmed. The levels drive the
SDK ``ToolAnnotations`` exposed in ``tools/list`` as well as any approval gate.

:class:`Role` maps human roles to the permission levels they may exercise:
viewers get read-only access, editors add standard (file-writing) operations,
and owners may run destructive tools too. The role matrix powers
:data:`ROLE_TOOL_MATRIX` and the RBAC middleware in :mod:`tianshang_scribe.mcp.transport`.
"""

from __future__ import annotations

from enum import Enum


class PermissionLevel(Enum):
    """Permission levels for MCP tools."""

    READ_ONLY = 'read_only'
    STANDARD = 'standard'
    DESTRUCTIVE = 'destructive'
    ADMIN = 'admin'


class Role(Enum):
    """Human/API-client roles with escalating tool access."""

    VIEWER = 'viewer'
    EDITOR = 'editor'
    OWNER = 'owner'


#: Permission level for every registered tool.
TOOL_PERMISSIONS: dict[str, PermissionLevel] = {
    'create_office_document': PermissionLevel.STANDARD,
    'edit_office_document': PermissionLevel.DESTRUCTIVE,
    'fill_template': PermissionLevel.STANDARD,
    'convert_document': PermissionLevel.STANDARD,
    'extract_document_data': PermissionLevel.READ_ONLY,
    'validate_template': PermissionLevel.READ_ONLY,
    'compare_documents': PermissionLevel.STANDARD,
    # Dedicated document-type-specific tools (v0.8.0 expansion).
    'create_excel_workbook': PermissionLevel.STANDARD,
    'edit_excel_workbook': PermissionLevel.DESTRUCTIVE,
    'create_presentation': PermissionLevel.STANDARD,
    'edit_presentation': PermissionLevel.DESTRUCTIVE,
    'analyze_excel_data': PermissionLevel.READ_ONLY,
    'compare_excel_workbooks': PermissionLevel.READ_ONLY,
    'extract_presentation_data': PermissionLevel.READ_ONLY,
}

#: Permission levels each role is allowed to exercise (escalating).
ROLE_PERMISSIONS: dict[Role, frozenset[PermissionLevel]] = {
    Role.VIEWER: frozenset({PermissionLevel.READ_ONLY}),
    Role.EDITOR: frozenset({PermissionLevel.READ_ONLY, PermissionLevel.STANDARD}),
    Role.OWNER: frozenset(
        {
            PermissionLevel.READ_ONLY,
            PermissionLevel.STANDARD,
            PermissionLevel.DESTRUCTIVE,
            PermissionLevel.ADMIN,
        }
    ),
}

#: Every registered tool each role may invoke (derived from ROLE_PERMISSIONS).
ROLE_TOOL_MATRIX: dict[Role, frozenset[str]] = {
    role: frozenset(
        tool for tool, level in TOOL_PERMISSIONS.items() if level in ROLE_PERMISSIONS[role]
    )
    for role in Role
}

#: Tools whose repeated identical calls produce identical results.
#: Read-only tools are idempotent by construction; file-writing tools are not
#: (each call re-writes, possibly creating temp output files), so they are
#: excluded even though the final bytes could in principle converge.
IDEMPOTENT_TOOLS: frozenset[str] = frozenset(
    {
        'extract_document_data',
        'validate_template',
        'analyze_excel_data',
        'compare_excel_workbooks',
        'extract_presentation_data',
    }
)


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

    Based on the explicit :data:`IDEMPOTENT_TOOLS` table (read-only tools are
    idempotent; file-writing tools are not).
    """
    return tool_name in IDEMPOTENT_TOOLS


def check_permission(tool_name: str, auto_approve: set[str]) -> bool:
    """Return whether ``tool_name`` may auto-execute.

    Tools are allowed when they are listed in ``auto_approve`` or when every
    sub-operation is read-only. Unknown tools default to ``STANDARD``.
    """
    if tool_name in auto_approve:
        return True
    return is_read_only(tool_name)


def roles_for(tool_name: str) -> set[Role]:
    """Return the set of roles allowed to invoke ``tool_name``."""
    level = levels_for(tool_name)
    return {role for role, levels in ROLE_PERMISSIONS.items() if level & set(levels)}


def role_allows(role: Role, tool_name: str) -> bool:
    """Return whether ``role`` may invoke ``tool_name``."""
    return tool_name in ROLE_TOOL_MATRIX.get(role, frozenset())


def parse_role(value: str | None) -> Role:
    """Coerce a string to a :class:`Role`, defaulting to ``OWNER``.

    Used by the RBAC middleware to interpret the ``X-Scribe-Role`` header.
    Unknown or empty values fall back to the most privileged role so existing
    deployments keep working unchanged.
    """
    if value:
        for role in Role:
            if role.value == value.strip().lower():
                return role
    return Role.OWNER
