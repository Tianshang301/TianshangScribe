"""File-system helpers shared across CLI, MCP server, and transformers."""

from __future__ import annotations

from pathlib import Path


def ensure_parent_dir(path: str | Path) -> None:
    """Create the parent directory of ``path`` if it does not exist yet."""
    parent = Path(path).parent
    if parent and not parent.exists():
        parent.mkdir(parents=True, exist_ok=True)


def check_overwrite(path: str | Path, force: bool) -> bool:
    """Return whether writing to ``path`` is allowed.

    Existing files require ``force=True`` unless the path does not exist.
    """
    return not (Path(path).exists() and not force)
