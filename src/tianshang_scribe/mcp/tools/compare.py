"""compare_documents — Compare two Word documents and manage snapshots.

The tool supports four sub-operations dispatched through ``options.action``
(a Discriminated-Union style aggregation to avoid tool proliferation):

- ``compare`` (default): paragraph-level diff of ``path_a`` vs ``path_b``.
- ``snapshot``: record the current paragraph state of ``path_a`` in a JSON
  snapshot store (default ``~/.tianshang-scribe/snapshots/``).
- ``list_snapshots``: list recorded snapshots for ``path_a``.
- ``restore``: write a snapshot back as a new .docx at ``path_b``.

Snapshots let an agent save a document's state before an edit and later
compare against or restore it.
"""

from __future__ import annotations

import hashlib
import json
import uuid
from pathlib import Path
from typing import Annotated, Any

from pydantic import Field

from tianshang_scribe.mcp.errors import McpErrorCode, error_response, success_response
from tianshang_scribe.mcp.schemas import ToolOptions

DEFAULT_SNAPSHOT_DIR = Path.home() / '.tianshang-scribe' / 'snapshots'


def _snapshot_store_dir(path: str, snapshot_dir: str | None) -> Path:
    """Return the snapshot store sub-directory for ``path`` (created lazily)."""
    base = Path(snapshot_dir) if snapshot_dir else DEFAULT_SNAPSHOT_DIR
    digest = hashlib.sha1(  # noqa: S324  # collision resistance irrelevant; used only as a store key
        Path(path).resolve().as_posix().encode()
    ).hexdigest()[:16]
    return base / digest


def snapshot_document(
    path: str,
    snapshot_dir: str | None = None,
) -> dict[str, Any]:
    """Record the current paragraph state of a Word document and return metadata."""
    if not Path(path).exists():
        return error_response(McpErrorCode.DOCUMENT_NOT_FOUND, f"'{path}' not found.")
    try:
        from tianshang_scribe.core.document import open_document

        engine = open_document(path)
        if not hasattr(engine, 'doc'):
            return error_response(
                McpErrorCode.UNSUPPORTED_FORMAT,
                'Snapshots are only supported for Word (.docx) files; '
                'Excel (.xlsx) and PowerPoint (.pptx) are not supported.',
            )
        paragraphs = [p.text for p in engine.doc.paragraphs]
        store_dir = _snapshot_store_dir(path, snapshot_dir)
        store_dir.mkdir(parents=True, exist_ok=True)
        snapshot_id = uuid.uuid4().hex[:16]
        source = Path(path)
        payload = {
            'id': snapshot_id,
            'source': str(source),
            'source_sha1': hashlib.sha1(source.read_bytes()).hexdigest(),  # noqa: S324  # non-cryptographic fingerprint
            'paragraphs': paragraphs,
            'paragraph_count': len(paragraphs),
        }
        (store_dir / f'{snapshot_id}.json').write_text(
            json.dumps(payload, ensure_ascii=False), encoding='utf-8'
        )
        return success_response(
            {
                'snapshot_id': snapshot_id,
                'source': str(source),
                'paragraph_count': len(paragraphs),
                'stored_at': str(store_dir / f'{snapshot_id}.json'),
            }
        )
    except Exception as e:
        return error_response(McpErrorCode.INTERNAL_ERROR, str(e))


def list_snapshots(path: str, snapshot_dir: str | None = None) -> dict[str, Any]:
    """Return the snapshots recorded for ``path``, newest first."""
    store_dir = _snapshot_store_dir(path, snapshot_dir)
    if not store_dir.exists():
        return success_response({'snapshots': []})
    snapshots: list[dict[str, Any]] = []
    for f in sorted(store_dir.glob('*.json'), reverse=True):
        try:
            data = json.loads(f.read_text(encoding='utf-8'))
            snapshots.append(
                {
                    'snapshot_id': data.get('id'),
                    'source': data.get('source'),
                    'paragraph_count': data.get('paragraph_count'),
                }
            )
        except (json.JSONDecodeError, OSError):
            continue
    return success_response({'snapshots': snapshots})


def restore_snapshot(
    path: str,
    snapshot_id: str,
    output_path: str,
    snapshot_dir: str | None = None,
) -> dict[str, Any]:
    """Write the snapshot identified by ``snapshot_id`` as a new .docx."""
    store_dir = _snapshot_store_dir(path, snapshot_dir)
    snapshot_file = store_dir / f'{snapshot_id}.json'
    if not snapshot_file.exists():
        return error_response(
            McpErrorCode.DOCUMENT_NOT_FOUND,
            f'Snapshot "{snapshot_id}" not found for "{path}".',
        )
    try:
        data = json.loads(snapshot_file.read_text(encoding='utf-8'))
        from tianshang_scribe.core.word_engine import WordEngine

        engine = WordEngine()
        engine.create()
        for text in data.get('paragraphs', []):
            if text:
                engine.add_text(text)
        engine.save(output_path)
        return success_response(
            {
                'restored_to': output_path,
                'snapshot_id': snapshot_id,
                'paragraph_count': len(data.get('paragraphs', [])),
            }
        )
    except Exception as e:
        return error_response(McpErrorCode.INTERNAL_ERROR, str(e))


def compare_documents(
    path_a: Annotated[str, Field(description='Path to the first document.')],
    path_b: Annotated[str, Field(description='Path to the second document.')],
    options: Annotated[ToolOptions | None, Field(description='Tool options.')] = None,
) -> dict[str, Any]:
    """Compare two Word documents, or snapshot/list/restore one (via options.action)."""
    action = (options.action if options else None) or 'compare'

    if action == 'snapshot':
        return snapshot_document(path_a, snapshot_dir=options.snapshot_dir if options else None)
    if action == 'list_snapshots':
        return list_snapshots(path_a, snapshot_dir=options.snapshot_dir if options else None)
    if action == 'restore':
        snapshot_id = options.snapshot_id if options else None
        if not snapshot_id:
            return error_response(
                McpErrorCode.INVALID_PARAMETER,
                'restore requires options.snapshot_id and path_b as the output target.',
            )
        return restore_snapshot(
            path_a,
            snapshot_id,
            path_b,
            snapshot_dir=options.snapshot_dir if options else None,
        )

    for p, label in [(path_a, 'path_a'), (path_b, 'path_b')]:
        if not Path(p).exists():
            return error_response(
                McpErrorCode.DOCUMENT_NOT_FOUND,
                f"'{label}': '{p}' not found.",
            )

    try:
        from tianshang_scribe.core.document import open_document

        engine_a = open_document(path_a)
        engine_b = open_document(path_b)

        if not hasattr(engine_a, 'doc') or not hasattr(engine_b, 'doc'):
            return error_response(
                McpErrorCode.UNSUPPORTED_FORMAT,
                'Document comparison is only supported for Word (.docx) files; '
                'Excel (.xlsx) and PowerPoint (.pptx) comparison is not yet available.',
            )

        paras_a = [p.text for p in engine_a.doc.paragraphs]
        paras_b = [p.text for p in engine_b.doc.paragraphs]

        added: list[dict[str, Any]] = []
        removed: list[dict[str, Any]] = []
        changed: list[dict[str, Any]] = []

        max_len = max(len(paras_a), len(paras_b))
        for i in range(max_len):
            text_a = paras_a[i] if i < len(paras_a) else ''
            text_b = paras_b[i] if i < len(paras_b) else ''
            if not text_a and text_b:
                added.append({'index': i, 'text': text_b[:200]})
            elif text_a and not text_b:
                removed.append({'index': i, 'text': text_a[:200]})
            elif text_a != text_b:
                changed.append(
                    {
                        'index': i,
                        'old': text_a[:200],
                        'new': text_b[:200],
                    }
                )

        return success_response(
            {
                'path_a': path_a,
                'path_b': path_b,
                'paragraphs_a': len(paras_a),
                'paragraphs_b': len(paras_b),
                'added': added,
                'removed': removed,
                'changed': changed,
                'identical': (len(added) == 0 and len(removed) == 0 and len(changed) == 0),
            }
        )
    except Exception as e:
        return error_response(McpErrorCode.INTERNAL_ERROR, str(e))
