# TianshangScribe MCP Server

> [中文版](./README.zh-CN.md)

MCP (Model Context Protocol) server for Office document processing. Enables AI Agents to **create**, **edit**, **fill templates**, **convert**, and **extract data** from Word, Excel, and PowerPoint documents — with native LaTeX-style formatting and mathematical formula support.

## Quick Start

```bash
# Install
git clone https://github.com/Tianshang301/TianshangScribe.git
cd TianshangScribe
pip install -e ".[dev]"

# Test
python mcp/test_server.py

# Use with Claude Code / Cursor
# Add to your MCP config:
```

```json
{
  "mcpServers": {
    "tianshang-scribe": {
      "command": "python",
      "args": ["-m", "mcp.server"],
      "cwd": "/path/to/TianshangScribe"
    }
  }
}
```

## Tools

### 1. `create_office_document`

Create Word, Excel, or PowerPoint documents with structured content.

```
Agent says:  "Generate a Q3 financial report"
Tool calls:  create_office_document(format="docx", content=[...])
```

**Parameters**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `format` | `"docx"` \| `"xlsx"` \| `"pptx"` | Yes | Output document format |
| `content` | `array` of ContentBlock | Yes | Ordered content blocks |
| `style` | `string` | No | Global style: `"font=Times,size=14,bold,color=FF0000"` |
| `template_data` | `object` | No | Key-value pairs to fill `{{placeholder}}` |
| `metadata` | `object` | No | Document properties: `{"title": "...", "author": "..."}` |
| `output_path` | `string` | No | Output file path (auto-generated if omitted) |
| `options` | `object` | No | `{"dry_run": true, "backup": true}` |

**ContentBlock types**

| `type` | Fields | Description |
|--------|--------|-------------|
| `paragraph` | `text`, `style` | Formatted text. Supports `\bfseries{}`, `$...$`, etc. |
| `heading` | `text`, `level` (1-6), `style` | Section heading |
| `formula` | `text` | LaTeX math formula — rendered as native OMML |
| `table` | `rows` (2D array) | Data table |
| `image` | `path` | Insert image from file |
| `page_break` | — | Page break |

**Example**

```json
{
  "format": "docx",
  "content": [
    {"type": "heading", "text": "Executive Summary", "level": 1},
    {"type": "paragraph", "text": "\\bfseries{Revenue:} \\color{0000FF}{$12.5M}. Growth: \\itshape{15.3%}."},
    {"type": "formula", "text": "\\sum_{i=1}^{n} x_i = \\frac{n(n+1)}{2}"},
    {"type": "table", "rows": [["Q1", "$3.5M"], ["Q2", "$4.2M"], ["Q3", "$4.8M"]]},
    {"type": "page_break"},
    {"type": "heading", "text": "Appendix", "level": 2}
  ],
  "metadata": {"title": "Q3 Report", "author": "AI Agent"}
}
```

### 2. `edit_office_document`

Modify an existing document with sequence of operations.

**Operation types**

| `action` | Key Fields | Description |
|----------|-----------|-------------|
| `replace` | `old_text`, `new_text`, `regex` | Find and replace text |
| `delete` | `target`, `regex` | Delete content |
| `modify` | `old_text`, `new_text` | Modify content (non-regex) |
| `style` | `style`, `apply_all` | Set style on all runs |
| `add` | `text`, `column` | Add text (Excel column support) |
| `clear` | — | Clear cell content |

```json
{
  "input_path": "report.docx",
  "operations": [
    {"action": "replace", "old_text": "2025", "new_text": "2026"},
    {"action": "style", "style": "font=Times New Roman,size=12", "apply_all": true}
  ]
}
```

### 3. `fill_template`

Fill placeholders in a template document with structured data. Supports nested keys (`{{user.name}}`) and loops (`{{#each items}}...{{/each}}`).

```json
{
  "template_path": "invitation.docx",
  "data": {
    "name": "Dr. Smith",
    "event": "AI Summit",
    "date": "2026-09-15"
  }
}
```

### 4. `convert_document`

Convert between formats.

| From | To | Supported |
|------|----|-----------|
| `docx` | `pdf`, `md`, `html` | All |
| `xlsx` | `pdf`, `csv`, `json`, `html` | All |
| `pptx` | `pdf` | Yes |

```json
{
  "input_path": "report.docx",
  "target_format": "pdf",
  "output_path": "report.pdf"
}
```

### 5. `extract_document_data`

Extract content from a document.

| `mode` | Returns |
|--------|---------|
| `metadata` | Author, title, subject, keywords |
| `text` | Full text content with block count |
| `structure` | Paragraphs/sections (Word), sheets (Excel), slides (PPT) |

```json
{
  "input_path": "report.docx",
  "mode": "text"
}
```

## LaTeX Markup Reference

All tools accept LaTeX-style markup in `text` fields:

| Markup | Effect |
|--------|--------|
| `\bfseries{text}` | **Bold** |
| `\itshape{text}` | *Italic* |
| `\underline{text}` | Underline |
| `\scshape{text}` | Small Caps |
| `\color{FF0000}{text}` | Colored text |
| `\fontsize{24}{text}` | Font size (pt) |
| `\fontfamily{Arial}{text}` | Specific font |
| `\heading{N}{text}` | Heading level N |
| `\newpage` | Page break |
| `\centering{...}` | Center align |
| `$E=mc^2$` | Inline math |
| `$$x=1$$` | Display math |

## Error Handling

All tools return structured error responses:

```json
{
  "success": false,
  "error_code": 1002,
  "error_message": "The document is password-protected.",
  "suggested_fix": "Provide the password or unlock the document first.",
  "retryable": true
}
```

**Error codes**

| Code | Name | Retryable |
|------|------|-----------|
| 0 | SUCCESS | — |
| 1001 | DOCUMENT_NOT_FOUND | No |
| 1002 | DOCUMENT_LOCKED | Yes |
| 1003 | UNSUPPORTED_FORMAT | No |
| 1004 | TEMPLATE_ERROR | Yes |
| 1005 | CONVERSION_FAILED | Yes |
| 1006 | INVALID_PARAMETER | No |
| 9999 | INTERNAL_ERROR | No |

## Dry Run & Backup

All tools support `options`:

```json
{
  "options": {
    "dry_run": true,
    "backup": true,
    "deterministic_id": "uuid-v4"
  }
}
```

- `dry_run`: Preview planned changes without writing files
- `backup`: Create `.bak` copy before modifying
- `deterministic_id`: Trackable operation ID for audit

## Architecture

```
MCP Client (Claude Code / Cursor / Agent)
    ↕ stdio JSON-RPC 2.0
mcp/server.py          ← Protocol dispatch
    ↕
mcp/tools/
    ├── create.py      ← WordEngine / ExcelEngine / PptEngine
    ├── edit.py        ← replace_text / set_style / clear_content
    ├── template.py    ← TemplateEngine
    └── convert.py     ← pdf.py / export methods
    ↕
src/core/              ← Document engines (python-docx, openpyxl, python-pptx)
```

- **Zero external MCP dependency** — pure stdio JSON-RPC 2.0
- **Transport**: stdio (Phase 1), SSE planned (Phase 2)
- **Protocol**: MCP 2024-11-05

## License

Apache-2.0
