# TianshangScribe

> [中文版](./readme/README.zh-CN.md)

Cross-platform CLI Office document processing tool. Create, edit, template-fill, and convert Word, Excel, and PowerPoint documents — with a built-in LaTeX style markup engine and native math formula renderer.

## Install

```bash
git clone https://github.com/Tianshang301/TianshangScribe.git
cd TianshangScribe
pip install -e ".[dev]"
```

Requires Python 3.10+ · python-docx · openpyxl · python-pptx · typer · rich · jinja2 · lxml

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

# Math formulas — auto-converted to native Word OMML
tianshang-scribe -w --create \
  --math "x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}" \
  --math "\sum_{i=0}^{n} i^2" \
  -o formulas.docx

# Template filling (JSON / CSV / YAML → {{placeholder}})
tianshang-scribe template.docx -t data.json -o filled.docx

# Convert to PDF
tianshang-scribe input.docx --topdf -o output.pdf

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
| `--clear` | Clear cell content | `--clear` |
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
| `\centering{...}` | Center align **†** |
| `\raggedright{...}` | Left align **†** |
| `\raggedleft{...}` | Right align **†** |
| `\linespread{1.5}{...}` | Line spacing **†** |
| `\indent{...}` / `\noindent{...}` | Indent **†** |
| `\heading{2}{Title}` | Insert heading |
| `\newpage` | Page break |
| `\includegraphics{path}` | Insert image |

**†** Paragraph-level formatting (creates a new paragraph).

### Font Configuration

| Command | Effect |
|---------|--------|
| `\setmainfont{Name}` | Default Western font |
| `\setCJKmainfont{Name}` | Default CJK font |
| `\setsansfont{Name}` | Sans-serif font |
| `\setmonofont{Name}` | Monospace font |

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
| Greek Letters | `\alpha` `\beta` `\gamma` … `\Gamma` `\Delta` `\Theta` … |
| Symbols | `\pm` `\times` `\div` `\cdot` `\infty` `\partial` `\nabla` `\forall` `\exists` … |
| Relations | `\leq` `\geq` `\neq` `\approx` `\equiv` `\propto` `\subset` `\supset` `\in` … |
| Arrows | `\to` `\rightarrow` `\leftarrow` `\mapsto` `\uparrow` … |
| Accents | `\hat{x}` `\bar{x}` `\tilde{x}` `\dot{x}` `\ddot{x}` `\vec{x}` `\widehat{x}` … |
| Brackets | `\left( \right)` `\left[ \right]` `\left\{ \right\}` |
| Math Fonts | `\mathrm{abc}` `\mathbf{abc}` `\mathit{abc}` `\mathcal{ABC}` `\mathbb{ABC}` `\mathsf{abc}` `\mathtt{abc}` |

### Math Typography

Conforms to mainstream math journal standards (AMS, Elsevier, Springer):

| Content | Style | Example |
|---------|-------|---------|
| Single-letter variables | *Italic* | `a` `b` `x` `y` |
| Digits | **Upright** | `0` `1` `2` … |
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
| `font` | `font_name` | Font name | Western font |
| `cjk-font` | `cjk_font_name` | Font name | CJK font |
| `size` | `font_size` | pt | Font size |
| `bold` | | flag | Bold |
| `italic` | | flag | Italic |
| `underline` | | flag | Underline |
| `color` | `font_color` | `FF0000` | Hex color |
| `align` | `alignment` | `left`/`center`/`right`/`justify` | Alignment |

Boolean keys (`bold` `italic` `underline`) are `True` when present.

## Template Filling

Supports JSON, CSV, and YAML data sources. Replaces `{{placeholder}}` in documents. Nested objects expand with dot notation. Loops iterate over list values.

```json
{
  "name": "John Doe",
  "date": "2026-07-28",
  "user": { "city": "Beijing" },
  "items": [
    { "product": "Widget", "price": "10" },
    { "product": "Gadget", "price": "20" }
  ]
}
```

```
{{name}}              →  John Doe
{{user.city}}         →  Beijing
{{#each items}}       →  repeats the block for each item
  {{product}}: {{price}}
{{/each}}
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
| Clear content | `--clear` |

## PPT Features

| Feature | Description |
|---------|-------------|
| Slide management | Add, delete, reorder slides (`--slide-add`, `--slide-delete`, `--slide-move`) |
| Layouts | Apply slide layouts by name or index (`--layout`) |
| Speaker notes | Add presenter notes (`--notes`) |
| Transitions | Set slide transitions — fade, push, wipe, etc. (`--transition`) |
| Export | Save slides as images (`--toimg`), convert to PDF (`--topdf`) |
| Protection | Set/remove password protection (`--protect`, `--unprotect`) |

## Exit Codes

| Code | Meaning |
|------|---------|
| `0` | Success |
| `1` | General error |
| `2` | Argument error |
| `3` | Not implemented |

## Architecture

```
src/
├── cli/               # Typer CLI entry
│   ├── main.py        # Command parsing & dispatch
│   └── global_opts.py # File path / type inference
├── core/              # Document engine abstraction
│   ├── document.py    # DocumentABC unified interface
│   ├── word_engine.py # Word engine (python-docx)
│   ├── excel_engine.py# Excel engine (openpyxl)
│   └── ppt_engine.py  # PPT engine (python-pptx)
├── rendering/         # Style & formula rendering
│   ├── styles.py      # TextStyle dataclass
│   ├── latex_parser.py # LaTeX markup parser
│   ├── math_omml.py   # LaTeX → OMML math converter
│   └── template.py    # Template filling engine
├── transform/         # Format conversion
│   └── pdf.py         # PDF export (LibreOffice)
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
| Math | Custom recursive descent parser → OMML XML |
| Templates | Jinja2 + docxtpl |
| PDF | LibreOffice headless |
| Quality | pytest (156 tests) · ruff · mypy |

## Development

```bash
git clone https://github.com/Tianshang301/TianshangScribe.git
cd TianshangScribe
pip install -e ".[dev]"

pytest tests/ -v        # Run tests
ruff check src/ tests/  # Lint
mypy src/               # Type check
```

## License

Apache-2.0
