from __future__ import annotations

from typing import Any

from src.core.document import DocumentABC


class TemplateEngine:

    def __init__(self, data_path: str) -> None:
        self._data_path = data_path
        self._data: dict[str, Any] = {}
        self._load()

    def _load(self) -> None:
        import csv
        import json
        from pathlib import Path

        path = Path(self._data_path)

        if path.suffix.lower() in ('.json',):
            with open(path, 'r', encoding='utf-8') as f:
                self._data = json.load(f)
        elif path.suffix.lower() in ('.csv',):
            with open(path, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                rows = list(reader)
                self._data = {'rows': rows, 'count': len(rows)}
        elif path.suffix.lower() in ('.yaml', '.yml'):
            try:
                import yaml
                with open(path, 'r', encoding='utf-8') as f:
                    self._data = defaultdict_from_yaml(yaml.safe_load(f))
            except ImportError:
                raise ImportError(
                    'PyYAML is required for YAML templates. '
                    'Install with: pip install pyyaml'
                )
        else:
            raise ValueError(f'Unsupported template format: {path.suffix}')

    def fill(self, engine: DocumentABC) -> int:
        _ = engine.doc.element.xml if hasattr(engine, 'doc') else ''
        count = 0
        for key, value in self._flatten_data(self._data).items():
            placeholder = f'{{{{{key}}}}}'
            if hasattr(engine, 'replace_text'):
                count += engine.replace_text(placeholder, str(value))
        return count

    def _flatten_data(self, data: dict[str, Any], prefix: str = '') -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in data.items():
            full_key = f'{prefix}.{key}' if prefix else key
            if isinstance(value, dict):
                result.update(self._flatten_data(value, full_key))
            elif isinstance(value, list):
                pass
            else:
                result[full_key] = value
        return result


def defaultdict_from_yaml(data: Any) -> Any:
    if isinstance(data, dict):
        return {k: defaultdict_from_yaml(v) for k, v in data.items()}
    elif isinstance(data, list):
        return [defaultdict_from_yaml(v) for v in data]
    return data
