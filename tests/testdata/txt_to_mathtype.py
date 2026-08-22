"""Convert the ``math_paper.txt`` test case into a Word file with native math.

Reads ``tests/testdata/math_paper.txt`` (Markdown-flavoured) using this tool's
own Word engine and renders it to a ``.docx`` whose formulas are Word-native
OMML (rendered by Word and LibreOffice alike, no MathType required):

- ``#`` / ``##`` lines  -> headings (``add_heading``)
- ``$$...$$`` blocks    -> block-level OMML equations (``add_math_formula``)
- ``$...$`` inline math -> inline OMML equations (``_add_omath``)
- ``**bold**`` / ``*italic*`` -> styled runs
- other lines           -> plain paragraphs

Usage:
    python tests/testdata/txt_to_mathtype.py [output.docx]
"""

from __future__ import annotations

import re
import sys
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from tianshang_scribe.core.word_engine import WordEngine

_BLOCK_MATH = re.compile(r'^\$\$(.+?)\$\$\s*$', re.S)
_INLINE_MATH = re.compile(r'\$(.+?)\$(?!\$)')
_STYLE_TOKEN = re.compile(r'(\*\*.+?\*\*|\*(?!\*).+?(?<!\*)\*)')


def _split_inline_math(text: str) -> list[tuple[str, bool]]:
    """Split ``text`` into (segment, is_math) pieces at ``$...$`` boundaries."""
    pieces: list[tuple[str, bool]] = []
    pos = 0
    for m in _INLINE_MATH.finditer(text):
        if m.start() > pos:
            pieces.append((text[pos : m.start()], False))
        pieces.append((m.group(1), True))
        pos = m.end()
    if pos < len(text):
        pieces.append((text[pos:], False))
    return pieces


def _append_formatted(paragraph: Any, text: str) -> None:
    """Append ``text`` to ``paragraph`` honouring ``**bold**``/``*italic*``."""
    pos = 0
    for m in _STYLE_TOKEN.finditer(text):
        if m.start() > pos:
            paragraph.add_run(text[pos : m.start()])
        token = m.group(0)
        if token.startswith('**'):
            run = paragraph.add_run(token[2:-2])
            run.bold = True
        else:
            run = paragraph.add_run(token[1:-1])
            run.italic = True
        pos = m.end()
    if pos < len(text):
        paragraph.add_run(text[pos:])


def _append_content(engine: WordEngine, paragraph: Any, text: str) -> None:
    """Append text with inline math to a paragraph, keeping segment order."""
    for segment, is_math in _split_inline_math(text):
        if is_math:
            engine._add_omath(paragraph, segment.strip())
        elif segment:
            _append_formatted(paragraph, segment)


def convert(txt_path: Path, out_path: Path) -> int:
    """Render the paper, returning the number of embedded OMML equations."""
    lines = txt_path.read_text(encoding='utf-8').splitlines()
    engine = WordEngine()
    engine.create()
    count = 0
    for raw in lines:
        line = raw.rstrip()
        if not line.strip():
            continue
        block = _BLOCK_MATH.match(line)
        if block:
            engine.add_math_formula(block.group(1).strip())
            count += 1
            continue
        if line.startswith('## '):
            engine.add_heading(line[3:].strip(), level=2)
            continue
        if line.startswith('# '):
            engine.add_heading(line[2:].strip(), level=1)
            continue
        if '$$' in line or '$' in line:
            paragraph = engine.doc.add_paragraph()
            _append_content(engine, paragraph, line)
            count += line.count('$') // 2
            continue
        paragraph = engine.doc.add_paragraph()
        _append_formatted(paragraph, line)
    engine.save(out_path)
    return count


def main() -> None:
    root = Path(__file__).resolve().parents[2]
    txt = root / 'tests' / 'testdata' / 'math_paper.txt'
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else txt.with_suffix('.mathtype.docx')
    count = convert(txt, out)
    print(f'Saved Word document (native OMML math): {out}')
    print(f'Embedded OMML equations: {count}')


if __name__ == '__main__':
    main()
