from __future__ import annotations

import re
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
            with open(path, encoding='utf-8-sig') as f:
                self._data = json.load(f)
        elif path.suffix.lower() in ('.csv',):
            with open(path, encoding='utf-8') as f:
                reader = csv.DictReader(f)
                rows = list(reader)
                self._data = {'rows': rows, 'count': len(rows)}
        elif path.suffix.lower() in ('.yaml', '.yml'):
            try:
                import yaml

                with open(path, encoding='utf-8') as f:
                    self._data = defaultdict_from_yaml(yaml.safe_load(f))
            except ImportError:
                raise ImportError(
                    'PyYAML is required for YAML templates. Install with: pip install pyyaml'
                ) from None
        else:
            raise ValueError(f'Unsupported template format: {path.suffix}')

    def fill(self, engine: DocumentABC) -> int:
        flat = self._flatten_data(self._data)

        if hasattr(engine, 'doc'):
            return self._fill_word(engine, flat)
        elif hasattr(engine, 'wb'):
            return self._fill_excel(engine, flat)
        else:
            return self._fill_simple(engine, flat)

    def _fill_word(self, engine: Any, flat: dict[str, Any]) -> int:
        count = 0
        doc = engine.doc

        i = 0
        while True:
            paragraphs = list(doc.paragraphs)
            if i >= len(paragraphs):
                break

            para = paragraphs[i]
            text = para.text.strip()
            each_match = re.match(r'\{\{#each\s+([^}\s]+)\}\}', text)
            if_match = re.match(r'\{\{#if\s+([^=}\s]+)(?:\s*=\s*([^}\s]+))?\}\}', text)
            unless_match = re.match(r'\{\{#unless\s+([^}\s]+)\}\}', text)

            if each_match:
                count += self._process_word_loop(doc, i, each_match.group(1), flat)
                i += 1
            elif if_match:
                count += self._process_word_if(
                    doc, i, if_match.group(1), if_match.group(2), flat, negate=False
                )
                i += 1
            elif unless_match:
                count += self._process_word_if(
                    doc, i, unless_match.group(1), None, flat, negate=True
                )
                i += 1
            else:
                i += 1

        for key, value in flat.items():
            placeholder = f'{{{{{key}}}}}'
            engine.replace_text(placeholder, str(value))
            count += 1

        return count

    def _process_word_if(
        self,
        doc: Any,
        start_idx: int,
        key: str,
        expected: str | None,
        flat: dict[str, Any],
        negate: bool,
    ) -> int:
        paragraphs = list(doc.paragraphs)
        end_idx = None
        for j in range(start_idx, len(paragraphs)):
            if '{{/if}}' in paragraphs[j].text or '{{/unless}}' in paragraphs[j].text:
                end_idx = j
                break

        if end_idx is None:
            return 0

        val = flat.get(key)
        condition = str(val) == expected if expected is not None else bool(val)

        if negate:
            condition = not condition

        for idx in range(start_idx, end_idx + 1):
            p = paragraphs[idx]
            if condition:
                raw = p.text
                cleaned = re.sub(
                    r'\{\{#(?:if|unless)\s+[^=}\s]+(?:\s*=\s*[^}\s]+)?\}\}',
                    '',
                    raw,
                )
                cleaned = cleaned.replace('{{/if}}', '').replace('{{/unless}}', '')
                if cleaned.strip():
                    p.text = cleaned
                else:
                    p.clear()
            else:
                p.clear()

        return 1

    def _process_word_loop(self, doc: Any, start_idx: int, key: str, flat: dict[str, Any]) -> int:
        items = flat.get(key)
        if not isinstance(items, list):
            return 0

        paragraphs = doc.paragraphs
        end_idx = None
        for j in range(start_idx + 1, len(paragraphs)):
            if '{{/each}}' in paragraphs[j].text:
                end_idx = j
                break

        if end_idx is None or end_idx <= start_idx:
            return 0

        template_paras = paragraphs[start_idx + 1 : end_idx]
        all_templates = []

        from copy import deepcopy

        from lxml import etree

        for item in items:
            for tp in template_paras:
                clone = deepcopy(tp._p)
                if isinstance(item, dict):
                    rendered = clone.xml
                    for ik, iv in item.items():
                        rendered = rendered.replace(f'{{{{{ik}}}}}', str(iv))
                    try:
                        new_el = etree.fromstring(rendered)
                    except Exception:
                        new_el = clone
                    all_templates.append(new_el)
                else:
                    text_el = etree.fromstring(clone.xml.replace('{{this}}', str(item)))
                    all_templates.append(text_el)

        anchor = paragraphs[end_idx]._p
        parent = anchor.getparent()
        for el in reversed(all_templates):
            parent.insert(parent.index(anchor) + 1, el)

        for j in range(end_idx, start_idx - 1, -1):
            p = paragraphs[j]._p
            p.getparent().remove(p)

        return len(items)

    def _fill_excel(self, engine: Any, flat: dict[str, Any]) -> int:
        ws = engine.wb.active
        max_row = ws.max_row or 1

        loop_keys = [k for k, v in flat.items() if isinstance(v, list)]

        for key, items in [(k, flat[k]) for k in loop_keys]:
            if not isinstance(items, list) or not items:
                continue

            template_row = None
            header_row = None
            for row_idx in range(1, max_row + 1):
                cell = ws.cell(row=row_idx, column=1)
                if cell.value and f'{{{{#each {key}}}}}' in str(cell.value):
                    header_row = row_idx
                    template_row = row_idx + 1
                    break

            if template_row is None:
                continue

            for item_idx, item in enumerate(items):
                copy_row = template_row + item_idx
                if copy_row != template_row:
                    ws.insert_rows(copy_row)
                for col_idx in range(1, ws.max_column + 1):
                    src = ws.cell(row=template_row, column=col_idx)
                    if src.value and isinstance(src.value, str):
                        new_val = src.value
                        if isinstance(item, dict):
                            for ik, iv in item.items():
                                new_val = new_val.replace(f'{{{{{ik}}}}}', str(iv))
                        else:
                            new_val = new_val.replace('{{this}}', str(item))
                        ws.cell(row=copy_row, column=col_idx, value=new_val)

            if header_row is not None:
                ws.cell(row=header_row, column=1, value='')
            ws.cell(row=template_row, column=1, value='')

            return len(items)

        flat_no_lists = {k: v for k, v in flat.items() if not isinstance(v, list)}
        return self._fill_simple(engine, flat_no_lists)

    def _fill_simple(self, engine: Any, flat: dict[str, Any]) -> int:
        count = 0
        for key, value in flat.items():
            if isinstance(value, list):
                continue
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
                result[full_key] = value
            else:
                result[full_key] = value
        return result


def defaultdict_from_yaml(data: Any) -> Any:
    if isinstance(data, dict):
        return {k: defaultdict_from_yaml(v) for k, v in data.items()}
    elif isinstance(data, list):
        return [defaultdict_from_yaml(v) for v in data]
    return data
