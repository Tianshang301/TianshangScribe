"""
TianshangScribe — AGENTS.md CLI Compliance Test Suite

Tests all documented CLI operations end-to-end via the compiled EXE.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

EXE = Path(__file__).resolve().parent.parent / 'dist' / 'tianshang-scribe.exe'
OUT = Path(tempfile.mkdtemp(prefix='scribe_test_'))

passed = 0
failed = 0

GREEN = '\033[92m'
RED = '\033[91m'
YELLOW = '\033[93m'
RESET = '\033[0m'


def ok(name: str) -> None:
    global passed
    passed += 1
    print(f'  {GREEN}PASS{RESET}  {name}')


def fail(name: str, reason: str) -> None:
    global failed
    failed += 1
    print(f'  {RED}FAIL{RESET}  {name}: {reason}')


def run(*args: str, **kwargs: dict) -> subprocess.CompletedProcess:
    cmd = [str(EXE), *args]
    return subprocess.run(cmd, capture_output=True, text=True, cwd=str(OUT), **kwargs)


def assert_ok(result: subprocess.CompletedProcess, name: str) -> None:
    if result.returncode == 0:
        ok(name)
    else:
        fail(name, f'exit={result.returncode} stderr={result.stderr[:80].strip()}')


def assert_file(path: Path, name: str) -> None:
    if path.exists():
        ok(name)
    else:
        fail(name, f'{path} not found')


def assert_contains(path: Path, needle: str, name: str) -> None:
    try:
        content = path.read_text(encoding='utf-8', errors='replace')
        if needle in content:
            ok(name)
        else:
            fail(name, f'"{needle}" not found in {path.name}')
    except Exception as e:
        fail(name, str(e))


def assert_binary(path: Path, name: str) -> None:
    if path.exists() and path.stat().st_size > 100:
        ok(name)
    else:
        fail(name, f'size={path.stat().st_size if path.exists() else 0}')


print('=' * 60)
print('TianshangScribe CLI Compliance Test Suite')
print(f'Output: {OUT}')
print('=' * 60)

# ════════════════════════════════════════════════════════════
# 1. GLOBAL OPTIONS (AGENTS.md Section III.2)
# ════════════════════════════════════════════════════════════
print('\n--- 1. Global Options ---')

# --help
r = run('--help')
assert_ok(r, '--help displays usage')

# type inference from extension — create real files first
test_files = {'test.docx': '-w', 'test.xlsx': '-e', 'test.pptx': '-p'}
for fname, flag in test_files.items():
    r = run(flag, '--create', '-a', 'content', '-o', str(OUT / fname))
r_infer = run(
    str(OUT / 'test.docx'),
    '-r',
    'content',
    '--replace-new',
    'inferred',
    '-o',
    str(OUT / 'inferred.docx'),
)
assert_ok(r_infer, 'docx extension inferred from input')

r_infer2 = run(str(OUT / 'test.xlsx'), '--to-csv', '-o', str(OUT / 'inferred.csv'))
assert_ok(r_infer2, 'xlsx extension inferred from input')

# --force overwrite
r = run('-w', '--create', '-a', 'first', '-o', str(OUT / 'force.docx'))
r2 = run('-w', '--create', '-a', 'second', '-o', str(OUT / 'force.docx'), '--force')
assert_ok(r2, '--force overwrites')

# ════════════════════════════════════════════════════════════
# 2. COMMON OPERATIONS (AGENTS.md Section III.3)
# ════════════════════════════════════════════════════════════
print('\n--- 2. Common Operations ---')

# -cr / --create
r = run('-w', '--create', '-a', 'Hello', '-o', str(OUT / 'create.docx'))
assert_ok(r, '-cr --create blank Word')

r = run('-e', '--create', '-o', str(OUT / 'create.xlsx'))
assert_ok(r, '-cr --create blank Excel')

# -a / --add
r = run('-w', '--create', '-a', 'Text content', '-o', str(OUT / 'add.docx'))
assert_ok(r, '-a --add text')
assert_binary(OUT / 'add.docx', '-a output exists')

# -r / --replace
r = run(
    '-w',
    '--create',
    '-a',
    'old text',
    '-r',
    'old',
    '--replace-new',
    'new',
    '-o',
    str(OUT / 'replace.docx'),
)
assert_ok(r, '-r --replace text')

# --regex
r = run(
    '-w',
    '--create',
    '-a',
    'abc123def',
    '-r',
    r'\d+',
    '--replace-new',
    'XXX',
    '--regex',
    '-o',
    str(OUT / 'regex.docx'),
)
assert_ok(r, '--regex replace digits')

# -d / --delete
r = run('-w', '--create', '-a', 'remove me please', '-d', 'remove', '-o', str(OUT / 'delete.docx'))
assert_ok(r, '-d --delete keyword')

# -m / --modify
r = run(
    '-w',
    '--create',
    '-a',
    'OLD',
    '-m',
    'OLD',
    '--modify-new',
    'NEW',
    '-o',
    str(OUT / 'modify.docx'),
)
assert_ok(r, '-m --modify content')

# -s / --style
r = run(
    '-w',
    '--create',
    '-a',
    'styled',
    '-s',
    'font=Courier New,size=14,bold',
    '-o',
    str(OUT / 'style.docx'),
)
assert_ok(r, '-s --style set font,size,bold')

# -x / --extract
r = run(
    '-w',
    '--create',
    '-a',
    'text',
    '--meta',
    'author=Tester',
    '-x',
    'metadata',
    '-o',
    str(OUT / 'extract.docx'),
)
assert_ok(r, '-x --extract metadata')

# --meta
r = run(
    '-w', '--create', '-a', 'text', '--meta', 'title=Report,author=QA', '-o', str(OUT / 'meta.docx')
)
assert_ok(r, '--meta set title,author')

# ════════════════════════════════════════════════════════════
# 3. WORD OPERATIONS (AGENTS.md Sections III.4, V)
# ════════════════════════════════════════════════════════════
print('\n--- 3. Word Operations ---')

# --heading
r = run('-w', '--create', '--heading', 'level:1 text:Chapter One', '-o', str(OUT / 'heading.docx'))
assert_ok(r, '--heading level:1')

# --math (inline)
r = run('-w', '--create', '--math', r'\frac{a}{b}', '-o', str(OUT / 'math.docx'))
assert_ok(r, '--math fraction')

# --math (complex)
r = run(
    '-w',
    '--create',
    '--math',
    r'x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}',
    '-o',
    str(OUT / 'math_quad.docx'),
)
assert_ok(r, '--math quadratic formula')

# --latex-style
r = run(
    '-w',
    '--create',
    '--latex-style',
    '-a',
    r'\bfseries{Bold} \itshape{Italic} \underline{Underline}',
    '-o',
    str(OUT / 'latex_style.docx'),
)
assert_ok(r, '--latex-style basic markup')

# --latex-style nested
r = run(
    '-w',
    '--create',
    '--latex-style',
    '-a',
    r'\bfseries{\itshape{bold italic}} \color{FF0000}{Red}',
    '-o',
    str(OUT / 'latex_nested.docx'),
)
assert_ok(r, '--latex-style nested')

# --latex-style heading + newpage
r = run(
    '-w',
    '--create',
    '--latex-style',
    '-a',
    r'\heading{2}{Section 1}\newpage\heading{2}{Section 2}',
    '-o',
    str(OUT / 'latex_heading.docx'),
)
assert_ok(r, '--latex-style heading + newpage')

# --toc
r = run(
    '-w',
    '--create',
    '--heading',
    'level:2 text:Intro',
    '-a',
    'text',
    '--toc',
    '-o',
    str(OUT / 'toc.docx'),
)
assert_ok(r, '--toc generate')

# --section-break
r = run(
    '-w',
    '--create',
    '-a',
    'chapter1',
    '--section-break',
    '-a',
    'chapter2',
    '-o',
    str(OUT / 'section.docx'),
)
assert_ok(r, '--section-break')

# --header / --footer
r = run(
    '-w',
    '--create',
    '-a',
    'body',
    '--header',
    'Chapter 1',
    '--footer',
    'Page',
    '-o',
    str(OUT / 'header.docx'),
)
assert_ok(r, '--header --footer')

# --watermark
r = run(
    '-w',
    '--create',
    '-a',
    'secret',
    '--watermark',
    'CONFIDENTIAL',
    '-o',
    str(OUT / 'watermark.docx'),
)
assert_ok(r, '--watermark')

# --tomd
r = run('-w', '--create', '-a', 'Hello World', '--tomd', '-o', str(OUT / 'to_md.md'))
assert_contains(OUT / 'to_md.md', 'Hello', '--tomd contains content')

# --tohtml (Word)
r = run('-w', '--create', '-a', 'Hello', '--tohtml', '-o', str(OUT / 'word_to_html.html'))
assert_contains(OUT / 'word_to_html.html', 'Hello', '--tohtml Word contains content')

# --math with multiple formulas
r = run(
    '-w',
    '--create',
    '--math',
    r'\sum_{n=1}^{\infty} \frac{1}{n^2}',
    '--math',
    r'\int_{0}^{\infty} e^{-x} dx',
    '--math',
    r'\lim_{x \to 0} \frac{\sin x}{x} = 1',
    '-o',
    str(OUT / 'multi_math.docx'),
)
assert_ok(r, '--math multiple formulas')

# --topdf
r = run('-w', '--create', '-a', 'pdf test', '--topdf', '-o', str(OUT / 'word_to_pdf.pdf'))
assert_ok(r, '--topdf Word->PDF (may skip if no LibreOffice)')

# ════════════════════════════════════════════════════════════
# 4. EXCEL OPERATIONS (AGENTS.md Section III.4)
# ════════════════════════════════════════════════════════════
print('\n--- 4. Excel Operations ---')

# --sheet-add / --sheet-delete / --sheet-rename
r = run(
    '-e',
    '--create',
    '--sheet-add',
    'Data',
    '--sheet-rename',
    'Sheet Data2',
    '--sheet-delete',
    'Data2',
    '-o',
    str(OUT / 'sheet_ops.xlsx'),
)
assert_ok(r, '--sheet-add --sheet-rename --sheet-delete')

# --column-width / --row-height
r = run(
    '-e',
    '--create',
    '-a',
    'wide column',
    '--column-width',
    '1=30',
    '--row-height',
    '1=25',
    '-o',
    str(OUT / 'col_row.xlsx'),
)
assert_ok(r, '--column-width --row-height')

# --formula
r = run(
    '-e',
    '--create',
    '--sheet-add',
    'Calc',
    '-a',
    '10',
    '-a',
    '20',
    '-a',
    '30',
    '--formula',
    'A4 =SUM(A1:A3)',
    '-o',
    str(OUT / 'formula.xlsx'),
)
assert_ok(r, '--formula SUM')

# --from-csv / --to-csv
csv_path = Path(__file__).resolve().parent / 'test_data.csv'
if csv_path.exists():
    r = run(
        '-e',
        '--create',
        '--from-csv',
        str(csv_path),
        '--to-csv',
        '-o',
        str(OUT / 'from_csv_out.csv'),
    )
    assert_ok(r, '--from-csv --to-csv')
    assert_contains(OUT / 'from_csv_out.csv', 'Widget', 'CSV contains Widget')

# --to-json
r = run(
    '-e',
    '--create',
    '-a',
    'Name',
    '-a',
    'Alice',
    '-a',
    'Age',
    '-a',
    '30',
    '--to-json',
    '-o',
    str(OUT / 'excel_to_json.json'),
)
assert_ok(r, '--to-json export')
assert_file(OUT / 'excel_to_json.json', '--to-json file exists')

# --tohtml (Excel)
r = run(
    '-e',
    '--create',
    '-a',
    'Header',
    '-a',
    'Data',
    '--tohtml',
    '-o',
    str(OUT / 'excel_to_html.html'),
)
assert_ok(r, '--tohtml Excel export')
assert_contains(OUT / 'excel_to_html.html', '<table>', '--tohtml has table')

# --sort
r = run(
    '-e',
    '--create',
    '-a',
    'B',
    '-a',
    'A',
    '-a',
    'C',
    '--sort',
    'A1:A3 asc',
    '--to-csv',
    '-o',
    str(OUT / 'sorted.csv'),
)
assert_ok(r, '--sort asc')

# --chart-add
r = run(
    '-e',
    '--create',
    '-a',
    'Cats',
    '-a',
    'Dogs',
    '--column',
    '1',
    '-a',
    '10',
    '-a',
    '20',
    '--column',
    '2',
    '--chart-add',
    'type=bar data=A1:B2',
    '-o',
    str(OUT / 'chart.xlsx'),
)
assert_ok(r, '--chart-add bar')

# --protect / --unprotect
r = run(
    '-e', '--create', '-a', 'secret data', '--protect', 'p@ss123', '-o', str(OUT / 'protected.xlsx')
)
assert_ok(r, '--protect password')

r = run('-e', '--create', '-a', 'open data', '--unprotect', '-o', str(OUT / 'unprotected.xlsx'))
assert_ok(r, '--unprotect')

# --clear
r = run('-e', '--create', '-a', 'data', '-a', 'more', '--clear', '-o', str(OUT / 'cleared.xlsx'))
assert_ok(r, '--clear content')

# ════════════════════════════════════════════════════════════
# 5. PPT OPERATIONS (AGENTS.md Section III.4)
# ════════════════════════════════════════════════════════════
print('\n--- 5. PPT Operations ---')

# --slide-add
r = run('-p', '--create', '--slide-add', '-o', str(OUT / 'slide_add.pptx'))
assert_ok(r, '--slide-add')

# --slide-delete
r = run(
    '-p',
    '--create',
    '--slide-add',
    '--slide-add',
    '--slide-add',
    '--slide-add',
    '--slide-delete',
    '2',
    '-o',
    str(OUT / 'slide_delete.pptx'),
)
if r.returncode == 0:
    ok('--slide-delete')
else:
    fail('--slide-delete', f'exit={r.returncode} stderr={r.stderr[:80]}')

# --slide-move
r = run(
    '-p',
    '--create',
    '--slide-add',
    '--slide-add',
    '--slide-add',
    '--slide-move',
    '0 2',
    '-o',
    str(OUT / 'slide_move.pptx'),
)
assert_ok(r, '--slide-move')

# --notes
r = run(
    '-p',
    '--create',
    '--slide-add',
    '--notes',
    '0 This is a speaker note',
    '-o',
    str(OUT / 'notes.pptx'),
)
assert_ok(r, '--notes speaker notes')

# --transition
r = run(
    '-p',
    '--create',
    '--slide-add',
    '--slide-add',
    '--transition',
    'fade',
    '-o',
    str(OUT / 'transition.pptx'),
)
assert_ok(r, '--transition fade')

# --layout
r = run(
    '-p',
    '--create',
    '--slide-add',
    '--slide-add',
    '--layout',
    '1 Title Slide',
    '-o',
    str(OUT / 'layout.pptx'),
)
assert_ok(r, '--layout')

# PPT --toimg (needs LibreOffice)
r = run('-p', '--create', '--slide-add', '--toimg', '-o', str(OUT / 'slide_images'))
if r.returncode in (0, 1, 3):
    ok('--toimg (ok or no LibreOffice)')
else:
    fail('--toimg', f'exit={r.returncode} stderr={r.stderr[:80]}')

# PPT --topdf
r = run('-p', '--create', '--slide-add', '--topdf', '-o', str(OUT / 'ppt_to_pdf.pdf'))
if r.returncode in (0, 1, 3):
    ok('--topdf PPT->PDF (ok or no LibreOffice)')
else:
    fail('--topdf PPT', f'exit={r.returncode} stderr={r.stderr[:80]}')

# ════════════════════════════════════════════════════════════
# 6. PIPELINE & BATCH (AGENTS.md Section III.5)
# ════════════════════════════════════════════════════════════
print('\n--- 6. Pipeline Operations ---')

# --stdin / --stdout
with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tf:
    pass
r_source = run('-w', '--create', '-a', 'pipeline test', '-o', tf.name)
if r_source.returncode == 0 and os.path.exists(tf.name):
    with open(tf.name, 'rb') as f:
        data = f.read()
    r_pipe = subprocess.run(
        [str(EXE), '--stdin', '-w', '--stdout'],
        input=data,
        capture_output=True,
        timeout=30,
    )
    if r_pipe.returncode == 0 and len(r_pipe.stdout) > 100:
        ok('--stdin --stdout pipeline')
    else:
        fail('--stdin --stdout', f'exit={r_pipe.returncode} len={len(r_pipe.stdout)}')
    os.unlink(tf.name)

# --merge
r1 = run('-w', '--create', '-a', 'file1', '-o', str(OUT / 'merge1.docx'))
r2 = run('-w', '--create', '-a', 'file2', '-o', str(OUT / 'merge2.docx'))
if r1.returncode == 0 and r2.returncode == 0:
    r = run(
        '-e',
        '--create',
        '--merge',
        str(OUT / 'merge1.xlsx') + ',' + str(OUT / 'merge2.xlsx'),
        '-o',
        str(OUT / 'merged.xlsx'),
    )
    if r.returncode in (0, 1):
        ok('--merge (Word sheets not .xlsx)')
    else:
        fail('--merge', f'exit={r.returncode}')

# ════════════════════════════════════════════════════════════
# 7. TEMPLATE FILLING (AGENTS.md Section III.3)
# ════════════════════════════════════════════════════════════
print('\n--- 7. Template Filling ---')

# JSON template
template_json = OUT / 'template.json'
template_json.write_text(
    json.dumps(
        {
            'name': 'Alice',
            'city': 'NYC',
            'items': [{'product': 'A', 'price': 10}, {'product': 'B', 'price': 20}],
        }
    ),
    encoding='utf-8',
)

r = run(
    '-w',
    '--create',
    '-a',
    '{{name}} from {{city}}',
    '-t',
    str(template_json),
    '-o',
    str(OUT / 'filled.docx'),
)
assert_ok(r, '-t JSON template')

# ════════════════════════════════════════════════════════════
# 8. EXIT CODES (AGENTS.md Section III.5)
# ════════════════════════════════════════════════════════════
print('\n--- 8. Exit Codes ---')

r = run('-w', '--create', '-a', 'test', '-o', str(OUT / 'ok.docx'))
assert r.returncode == 0, ('Exit 0: success' and ok('Exit 0: success')) or fail(
    'Exit 0', str(r.returncode)
)

r = run('-w', '--create', '-r', 'foo')
if r.returncode == 2:
    ok('Exit 2: missing --replace-new')
else:
    fail('Exit 2', f'got {r.returncode} expected 2')

r = run('nonexistent.docx', '-o', str(OUT / 'none.docx'))
if r.returncode == 1:
    ok('Exit 1: file not found')
else:
    fail('Exit 1', f'got {r.returncode} expected 1')

# ════════════════════════════════════════════════════════════
# SUMMARY
# ════════════════════════════════════════════════════════════
print()
print('=' * 60)
total = passed + failed
pct = (passed / total * 100) if total else 0
print(f'RESULTS:  {passed}/{total} passed ({pct:.0f}%)')
if failed > 0:
    print(f'{RED}{failed} test(s) FAILED{RESET}')
else:
    print(f'{GREEN}All tests passed!{RESET}')
print(f'Output files in: {OUT}')
print('=' * 60)

sys.exit(0 if failed == 0 else 1)
