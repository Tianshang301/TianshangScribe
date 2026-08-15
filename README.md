# TianshangScribe

> [中文版](./readme/README.zh-CN.md)

<p align="center">
  <a href="https://glama.ai/mcp/servers/Tianshang301/TianshangScribe">
    <img alt="TianshangScribe MCP server"
         src="https://glama.ai/mcp/servers/Tianshang301/TianshangScribe/badges/card.svg"
         width="380">
  </a>
</p>

[![PyPI](https://img.shields.io/badge/pypi-tianshang--scribe-blue)](https://pypi.org/project/tianshang-scribe/)
[![CI](https://github.com/Tianshang301/TianshangScribe/actions/workflows/ci.yml/badge.svg)](https://github.com/Tianshang301/TianshangScribe/actions)
[![License](https://img.shields.io/badge/license-Apache%202.0-green)](LICENSE)
[![TianshangScribe MCP server](https://glama.ai/mcp/servers/Tianshang301/TianshangScribe/badges/score.svg)](https://glama.ai/mcp/servers/Tianshang301/TianshangScribe)

Cross-platform Office document processing for developers, CLI automation, and AI agents. Create, edit, template-fill, and convert Word (.docx), Excel (.xlsx), and PowerPoint (.pptx) documents, with LaTeX-style markup, native OMML math formulas, and a template engine ({{placeholders}}, {{#each}} loops, {{#if}} conditions). Ships an MCP Server with 7 tools (create, edit, fill template, convert, extract, validate, compare) over stdio, SSE, and Streamable HTTP transports, with bearer-token auth and rate limiting.

> **Warning: Unstable API \u2014 breaking changes expected**
>
> This project is pre-1.0 (0.x). The CLI options, MCP tool signatures, template
> syntax, and output formats are **not frozen** and may change without notice.
> **Compatibility commitment**: any breaking change will be announced in the
> CHANGELOG at least one release in advance and accompanied by a migration guide.
> For production use, pin to a specific version and review the CHANGELOG before
> upgrading.


## Install

```bash
pip install tianshang-scribe

# Or from source:
git clone https://github.com/Tianshang301/TianshangScribe.git
cd TianshangScribe
pip install -e ".[dev]"
```

### Linux Deployment

**Docker** (recommended for MCP Server over Streamable HTTP):
```bash
git clone https://github.com/Tianshang301/TianshangScribe.git
cd TianshangScribe
docker compose up -d
# Streamable HTTP MCP Server at http://localhost:8080/mcp
# (override transport / auth / rate limits via TIANSHANG_SCRIBE_* env vars)
```

**.deb package** (Debian / Ubuntu):
```bash
# Download from GitHub Releases
sudo dpkg -i tianshang-scribe_0.6.0_all.deb
tianshang-scribe --help
```

**pipx** (isolated CLI):
```bash
pipx install tianshang-scribe
tianshang-scribe --help
```

Requires Python 3.10+ · python-docx · openpyxl · python-pptx · typer · rich · lxml

## Quick Start

```bash
# Create a Word document
tianshang-scribe -w --create -a "Hello World" -o hello.docx

# Replace text (--regex for regex mode)
tianshang-scribe input.docx -r "old" --replace-new "new" -o output.docx

# LaTeX markup with nesting
tianshang-scribe -w --create --latex-style \
  -s "font=Times New Roman,size=14" \
  -a "\bfseries{\itshape{bold italic}} \fontsize{24}{Heading} \color{FF0000}{red}" \
  -o styled.docx

# Math formulas —auto-converted to native Word OMML
tianshang-scribe -w --create \
  --math "x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}" \
  --math "\sum_{i=0}^{n} i^2" \
  -o formulas.docx

# Template filling (JSON / CSV / YAML →{{placeholder}})
tianshang-scribe template.docx -t data.json -o filled.docx

# Convert to PDF (office2pdf ~2MB, or LibreOffice fallback)
tianshang-scribe input.docx --topdf -o output.pdf

# MCP Server —stdio mode (Claude Code / Cursor)
python -m tianshang_scribe.mcp.server

# MCP Server —SSE mode (Dify / Coze / FastGPT)
python -m tianshang_scribe.mcp.server --transport sse --port 8080

# Excel: import CSV, sort, export JSON
tianshang-scribe -e --create --from-csv data.csv --sort "A1:A10 asc" --to-json -o out.json

# Excel: add formula, protect workbook
tianshang-scribe budget.xlsx --formula "B10 =SUM(B2:B9)" --protect "p@ss" -o protected.xlsx
```

## Global Options

| Parameter | Description |
|-----------|-------------|
| `input_file` | Input document path (omit with `--create`) |
| `-w` `--word` | Process Word document |
| `-e` `--excel` | Process Excel workbook |
| `-p` `--ppt` | Process PowerPoint presentation |
| `-o` `--output` | Output file path |
| `--force` | Allow overwriting existing files |
| `--topdf` | Output as PDF |
| `--stdin` | Read from standard input |
| `--stdout` | Write to standard output |

When `-w/-e/-p` is omitted, the document type is inferred from the input file extension.

## Operations

| Option | Description | Example |
|--------|-------------|---------|
| `-cr` `--create` | Create blank document | `--create -w` |
| `-a` `--add` | Add text | `-a "Hello"` |
| `--column` | Target column for `--add` | `--column 2` |
| `-r` `--replace` | Find and replace | `-r "foo" --replace-new "bar"` |
| `-d` `--delete` | Delete content | `-d "keyword"` |
| `-cl` `--clear` | Clear content / formats / links | `--clear formats` |
| `-m` `--modify` | Modify content | `-m "old" --modify-new "new"` |
| `-s` `--style` | Set style | `-s "font=Times,size=14,bold"` |
| `-t` `--template` | Template filling | `-t data.json` |
| `-x` `--extract` | Extract data (`metadata`) | `-x metadata` |
| `--meta` | Set properties | `--meta "title=Report,author=John"` |
| `--latex-style` | Enable LaTeX parsing | |
| `--math` | Add math formula (Word) | `--math "\frac{a}{b}"` |
| `--heading` | Add heading (Word) | `--heading "level:1 text:Intro"` |
| `--regex` | Regex mode | Use with `--replace` `--delete` |
| `--merge` | Merge files | `--merge "a.docx,b.docx"` |
| `--split` | Split document | `--split by-page` |
| `--comment` | Add comment / notes | `--comment "Note text"` |
| `--add-table` | Add table (Word) | `--add-table "H1,H2\|a1,a2"` |
| `--chart-add` | Add chart (Excel) | `--chart-add "type=bar data=B1:C10"` |
| `--batch` | Batch mode | `--batch` |
| `--files` | Glob pattern for batch | `--files "reports/*.docx"` |
| `--schedule-db` | Schedule SQLite DB path | `--schedule-db ~/.tianshang-scribe/schedules.db` |
| `--schedule-add` | Register schedule | `--schedule-add "daily\|0 9 * * *\|echo hi"` |
| `--schedule-rm` | Remove schedule | `--schedule-rm daily` |
| `--schedule-list` | List schedules | `--schedule-list` |
| `--schedule-run` | Run schedule now | `--schedule-run daily` |
| `--schedule-run-all` | Run due schedules | `--schedule-run-all` |
| `--run-script` | Run script in sandbox | `--run-script build.py` |
| `--stdin` | Read from stdin | |
| `--stdout` | Write to stdout | |

## Word-Specific Options

| Option | Description | Example |
|--------|-------------|---------|
| `--heading` | Add heading | `--heading "level:1 text:Intro"` |
| `--math` | Add math formula | `--math "\frac{a}{b}"` |
| `--latex-style` | Enable LaTeX markup | |
| `--toc` | Generate table of contents | `--toc` |
| `--section-break` | Insert section break | `--section-break` |
| `--header` | Set page header | `--header "Chapter 1"` |
| `--footer` | Set page footer | `--footer "Page X"` |
| `--watermark` | Text watermark | `--watermark "DRAFT"` |
| `--tomd` | Convert to Markdown | `--tomd` |
| `--tohtml` | Convert to HTML | `--tohtml` |

## Excel-Specific Options

| Option | Description | Example |
|--------|-------------|---------|
| `--sheet-add` | Add worksheet | `--sheet-add "Q1"` |
| `--sheet-delete` | Delete worksheet | `--sheet-delete "Sheet2"` |
| `--sheet-rename` | Rename worksheet | `--sheet-rename "Old New"` |
| `--column-width` | Set column width | `--column-width "2=20"` |
| `--row-height` | Set row height | `--row-height "3=30"` |
| `--formula` | Set cell formula | `--formula "A1 =SUM(B1:B10)"` |
| `--from-csv` | Import CSV data | `--from-csv data.csv` |
| `--sort` | Sort range | `--sort "A1:A10 asc"` |
| `--chart-add` | Add chart | `--chart-add "type=bar data=B1:C10"` |
| `--protect` | Set password | `--protect "p@ss"` |
| `--unprotect` | Remove password | `--unprotect` |
| `--to-csv` | Export as CSV | |
| `--to-json` | Export as JSON | |
| `--to-html` | Export as HTML | |

## LaTeX Style Markup

Embed the following markup in `--add` content. Enable with `--latex-style`. Supports nesting.

| Syntax | Effect |
|--------|--------|
| `\bfseries{text}` | Bold |
| `\itshape{text}` | Italic |
| `\scshape{text}` | Small caps |
| `\underline{text}` | Underline |
| `\rmfamily{text}` | Roman (serif) |
| `\sffamily{text}` | Sans-serif |
| `\ttfamily{text}` | Monospace |
| `\fontfamily{Arial}{text}` | Specific font |
| `\fontsize{18}{text}` | Font size (pt) |
| `\color{FF0000}{text}` | Color (hex) |
| `\centering{...}` | Center align **—* |
| `\raggedright{...}` | Left align **—* |
| `\raggedleft{...}` | Right align **—* |
| `\linespread{1.5}{...}` | Line spacing **—* |
| `\indent{...}` / `\noindent{...}` | Indent **—* |
| `\heading{2}{Title}` | Insert heading |
| `\newpage` | Page break |
| `\includegraphics{path}` | Insert image |

**—* Paragraph-level formatting (creates a new paragraph).

### Font Configuration

| Command | Effect |
|---------|--------|
| `\setmainfont{Name}` | Default Western font |
| `\setCJKmainfont{Name}` | Default CJK font |
| `\setsansfont{Name}` | Sans-serif font |
| `\setCJKsansfont{Name}` | CJK sans-serif font |
| `\setmonofont{Name}` | Monospace font |
| `\setCJKmonofont{Name}` | CJK monospace font |

Word OOXML natively separates `w:ascii` (Western) and `w:eastAsia` (CJK) fonts, enabling automatic font switching in mixed-script text.

## Math Formulas

LaTeX math formulas via `--math` are converted to native Word OMML (Office Math Markup Language).

### Supported Syntax

| Category | Commands |
|----------|----------|
| Fractions | `\frac{num}{den}` |
| Roots | `\sqrt{content}` `\sqrt[n]{content}` |
| Sup/Subscripts | `x^{2}` `x_{i}` `x_{i}^{n}` |
| Sums/Integrals | `\sum` `\int` `\oint` `\prod` `\coprod` `\bigcup` `\bigcap` `\bigvee` `\bigwedge` |
| Limits | `\lim_{x \to 0}` `\max` `\min` `\sup` `\inf` |
| Named Functions | `\sin` `\cos` `\tan` `\cot` `\sec` `\csc` `\log` `\ln` `\det` `\Pr` `\gcd` `\deg` `\dim` `\hom` `\ker` `\arg` |
| Greek Letters | `\alpha` `\beta` `\gamma` —`\Gamma` `\Delta` `\Theta` —|
| Symbols | `\pm` `\times` `\div` `\cdot` `\infty` `\partial` `\nabla` `\forall` `\exists` —|
| Relations | `\leq` `\geq` `\neq` `\approx` `\equiv` `\propto` `\subset` `\supset` `\in` —|
| Arrows | `\to` `\rightarrow` `\leftarrow` `\mapsto` `\uparrow` —|
| Accents | `\hat{x}` `\bar{x}` `\tilde{x}` `\dot{x}` `\ddot{x}` `\vec{x}` `\widehat{x}` —|
| Brackets | `\left( \right)` `\left[ \right]` `\left\{ \right\}` |
| Math Fonts | `\mathrm{abc}` `\mathbf{abc}` `\mathit{abc}` `\mathcal{ABC}` `\mathbb{ABC}` `\mathsf{abc}` `\mathtt{abc}` |

### Math Typography

Conforms to mainstream math journal standards (AMS, Elsevier, Springer):

| Content | Style | Example |
|---------|-------|---------|
| Single-letter variables | *Italic* | `a` `b` `x` `y` |
| Digits | **Upright** | `0` `1` `2` —|
| Named functions | **Upright** | `\sin` `\cos` `\log` |
| Lowercase Greek | *Italic* | `\alpha` `\beta` `\gamma` |
| Uppercase Greek | **Upright** | `\Gamma` `\Delta` `\Theta` |

### Auto-Detection

Commands in `--add` text are automatically recognized as math even without `$...$` wrapping:
- With arguments: `\frac` `\sqrt` `\sum` `\int` `\prod` `\lim`
- Accents: `\hat{x}` `\bar{x}` `\vec{x}` etc.
- Unary operators: `\sin` `\cos` `\tan` `\log` `\ln` etc.
- `H_{2}O` and `m^{2}` in plain text become Unicode sub/superscripts (H₂O / m²)

## Style Syntax

`--style` uses comma-separated key-value pairs:

```bash
--style "font=Times New Roman,size=14,bold,italic,color=FF0000,align=center"
```

| Key | Aliases | Value | Description |
|-----|---------|-------|-------------|
| `font` | `font_name`, `font-family` | Font name | Western font |
| `cjk-font` | `cjk_font_name`, `cjk-font-family` | Font name | CJK font |
| `size` | `font_size`, `font-size` | pt | Font size |
| `bold` | | flag | Bold |
| `italic` | | flag | Italic |
| `underline` | | flag | Underline |
| `color` | `font_color`, `font-color` | `FF0000` | Hex color |
| `align` | `alignment` | `left`/`center`/`right`/`justify` | Alignment |

Boolean keys (`bold` `italic` `underline`) are `True` when present.

## Template Filling

Supports JSON, CSV, and YAML data sources. Replaces `{{placeholder}}` in documents. Nested objects expand with dot notation. Loops iterate over list values. Conditionals show/hide blocks.

```json
{
  "name": "John Doe",
  "date": "2026-07-28",
  "user": { "city": "Beijing" },
  "show": true,
  "paid": false,
  "items": [
    { "product": "Widget", "price": "10" },
    { "product": "Gadget", "price": "20" }
  ]
}
```

```
{{name}}              → John Doe
{{user.city}}         → Beijing
{{#each items}}       → repeats the block for each item
  {{product}}: {{price}}
{{/each}}
{{#if show}}          → shown only when show is truthy
  Confidential content
{{/if}}
{{#if role=admin}}    → shown only when role equals "admin"
  Admin dashboard
{{/if}}
{{#unless paid}}      → shown only when paid is falsy
  Payment required
{{/unless}}
```

## Excel Features

| Feature | CLI Option |
|---------|------------|
| Sheet management | `--sheet-add` `--sheet-delete` `--sheet-rename` |
| Column/row sizing | `--column-width` `--row-height` |
| Formulas | `--formula "A1 =SUM(B1:B10)"` |
| Data import | `--from-csv` |
| Data export | `--to-csv` `--to-json` `--to-html` |
| Sorting | `--sort "A1:A10 asc"` |
| Charts | `--chart-add "type=bar data=B1:C10"` |
| Protection | `--protect` `--unprotect` |

## PPT Features

| Feature | Description |
|---------|-------------|
| Slide management | Add, delete, reorder slides (`--slide-add`, `--slide-delete`, `--slide-move`) |
| Layouts | Apply slide layouts by name or index (`--layout`) |
| Speaker notes | Add presenter notes (`--notes`) |
| Math formulas | `$...$` / `$$...$$` rendered as native OMML |
| Transitions | Set slide transitions —fade, push, wipe, etc. (`--transition`) |
| Export | Save slides as images (`--toimg`), convert to PDF (`--topdf`) |
| Media compression | Compress images (`--compress-media "1920,80"`) |
| Protection | Set/remove password (`--protect`, `--unprotect`) |

## Exit Codes

| Code | Meaning |
|------|---------|
| `0` | Success |
| `1` | General error |
| `2` | Argument error |
| `3` | Not implemented |

## MCP Server

TianshangScribe includes an MCP (Model Context Protocol) server —AI Agents can create, edit, fill templates, convert, and extract data from Office documents.

### Quick Connect

**stdio** (Claude Code, Cursor):
```json
{"mcpServers": {"tianshang-scribe": {
  "command": "python", "args": ["-m", "tianshang_scribe.mcp.server"]
}}}
```

**SSE** (Dify, Coze, FastGPT):
```bash
python -m tianshang_scribe.mcp.server --transport sse --host 0.0.0.0 --port 8080
```
```json
{"mcpServers": {"tianshang-scribe": {
  "url": "http://localhost:8080/sse", "transport": "sse"
}}}
```

### Tools (7)

| Tool | Description |
|------|-------------|
| `create_office_document` | Create .docx / .xlsx / .pptx with structured content blocks |
| `edit_office_document` | Replace, delete, modify, style, add operations on existing docs |
| `fill_template` | Fill `{{placeholders}}` with data; supports `{{#each}}` / `{{#if}}` |
| `convert_document` | Convert between formats (docx↔pdf/md/html, xlsx↔csv/json) |
| `extract_document_data` | Extract metadata, full text, or document structure |
| `validate_template` | Pre-check template placeholders against data before filling |
| `compare_documents` | Paragraph-level diff between two .docx files |

### Capabilities

| Feature | Detail |
|---------|--------|
| **Protocol** | MCP 2024-11-05 · stdio + SSE · JSON-RPC 2.0 |
| **Resources** | `resources/list` + `resources/read` —documents exposed as readable URIs |
| **Prompts** | 5 built-in workflow templates (`prompts/list` + `prompts/get`) |
| **Progress** | `notifications/progress` during PDF conversion and long operations |
| **Response** | Multi-type `content[]`: text message + resource (file URI, MIME type, size) |
| **Schema** | `enum`, `default`, `examples`, `minimum/maximum` constraints on all params |

### Production (SSE only)

```bash
# With authentication
TIANSHANG_SCRIBE_AUTH_TOKEN="secret" \
python -m tianshang_scribe.mcp.server --transport sse --host 0.0.0.0 --port 8080

# Health check
curl http://localhost:8080/health
# {"status":"ok","version":"0.6.0","uptime_seconds":3600,"active_sessions":3,"tools_available":7}

# CORS whitelist
python -m tianshang_scribe.mcp.server --transport sse --cors-origins "https://coze.com,https://dify.ai"
```

**Endpoints**: `GET /health` · `GET /sse` · `POST /message?session_id=X`

Full documentation: [docs/mcp/README.md](docs/mcp/README.md).

```bash
python tests/integration/mcp/mcp_stdio_smoke.py     # 9/9 quick tests (stdio)
python tests/integration/mcp/test_sse.py        # 3/3 SSE transport tests
python tests/integration/mcp/mcp_agent_sim.py      # 11-scenario Agent simulation
```

## Architecture

```
src/
└── tianshang_scribe/    # importable package (tianshang_scribe.*)
    ├── cli/               # Typer CLI entry
    │  ├── main.py        # Command parsing & dispatch
    │  └── global_opts.py # File path / type inference
    ├── core/              # Document engine abstraction
    │  ├── document.py    # DocumentABC unified interface
    │  ├── word_engine.py # Word engine (python-docx)
    │  ├── excel_engine.py# Excel engine (openpyxl)
    │  └── ppt_engine.py  # PPT engine (python-pptx)
    ├── rendering/         # Style & formula rendering
    │  ├── styles.py      # TextStyle dataclass
    │  ├── latex_parser.py # LaTeX markup parser
    │  ├── math_omml.py   # LaTeX →OMML math converter
    │  └── template.py    # Template filling engine
    ├── transform/         # Format conversion
    │  └── pdf.py         # PDF export (office2pdf + LibreOffice)
    ├── mcp/                    # MCP Server (official mcp SDK 2.x)
    │  ├── server.py           # build_server + entry (stdio / SSE / Streamable HTTP)
    │  ├── transport.py        # transport wiring + ASGI middleware
    │  ├── schemas.py          # pydantic models + as_dict
    │  ├── auth.py             # Bearer token auth
    │  ├── rate_limit.py       # token bucket rate limiting
    │  ├── metrics.py          # Prometheus-style metrics
    │  ├── security.py         # read-only / destructive classification
    │  ├── prompts.py          # 5 prompt workflows
    │  ├── tools/              # 7 Agent tools
    │  │  ├── _registry.py    # tool registry (schemas auto-derived)
    │  │  ├── create.py / edit.py / template.py / convert.py
    │  │  ├── validate.py / compare.py
    │  └── errors.py           # structured error codes + fixes
    └── utils/             # Utility functions
        └── file_utils.py
```

### Tech Stack

| Component | Technology |
|-----------|------------|
| CLI | Typer + Rich |
| Word | python-docx |
| Excel | openpyxl |
| PPT | python-pptx |
| Math | Custom recursive descent parser →OMML XML |
| Templates | Custom engine ({{placeholder}}, {{#each}}, {{#if}}) |
| PDF | office2pdf (~2MB Rust binary, zero deps) + LibreOffice fallback |
| Quality | pytest (414 tests) · ruff · mypy |

## Build EXE

```bash
pip install pyinstaller
pyinstaller --onefile --name tianshang-scribe --hidden-import openpyxl.cell._writer --hidden-import openpyxl.cell.read_only --hidden-import openpyxl.styles --hidden-import openpyxl.chart --hidden-import openpyxl.comments src/tianshang_scribe/cli/main.py
# dist/tianshang-scribe.exe (~35 MB)
```

## Demo

```bash
python -m demo.generate_demos
# demo/demo_word.docx   —LaTeX + math + TOC + watermark
# demo/demo_excel.xlsx  —CSV import + formulas + chart + protection
# demo/demo_ppt.pptx    —slides + notes + transitions + math formulas
```

CLI compliance test:

```bash
python demo/test_cli.py
```

## Development

```bash
git clone https://github.com/Tianshang301/TianshangScribe.git
cd TianshangScribe
pip install -e ".[dev]"

pytest tests/ -v        # Run tests
ruff check src/tianshang_scribe/ tests/  # Lint
mypy src/tianshang_scribe/               # Type check
```

## License

Apache-2.0
