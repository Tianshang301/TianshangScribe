"""One-shot CLI command for the tianshang-scribe application."""

from __future__ import annotations

from pathlib import Path
from typing import Annotated, Any

import click.utils as _click_utils
import typer
import typer._click.utils as _typer_click_utils
from rich.console import Console

from src.cli.global_opts import (
    check_overwrite,
    determine_output_path,
    is_reverse_source,
    parse_table_input,
    resolve_doc_type,
)
from src.core.document import DocumentType, create_document, detect_document_type, open_document
from src.rendering.template import TemplateEngine


def _no_expand(args, **kwargs):
    """Disable Click's Windows argv glob expansion (see module docstring)."""
    return list(args)


# Click 8+ / Typer expand glob patterns in argv on Windows by default
# (``windows_expand_args``). The CLI controls its own globbing via ``--files``
# / positional patterns, so disable Click's expansion to let patterns through
# literally. Typer vendors its own ``_click`` copy, so both are patched.
_click_utils._expand_args = _no_expand
_typer_click_utils._expand_args = _no_expand

app = typer.Typer(
    name='tianshang-scribe',
    help='天殇·书契 — Cross-platform CLI Office document processing tool',
)
console = Console()

# Shared option help strings — single source of truth used by BOTH the one-shot
# command and the ``open`` subcommand so their CLI parameters stay in sync.
LATEX_STYLE_HELP = 'Enable LaTeX style markup parsing'
WORD_HELP = 'Process Word document'
EXCEL_HELP = 'Process Excel workbook'
PPT_HELP = 'Process PowerPoint presentation'


@app.command()
def main(
    input_file: Annotated[
        str | None,
        typer.Argument(help='Input document path (omit with --create)'),
    ] = None,
    word: Annotated[
        bool,
        typer.Option('-w', '--word', help=WORD_HELP),
    ] = False,
    excel: Annotated[
        bool,
        typer.Option('-e', '--excel', help=EXCEL_HELP),
    ] = False,
    ppt: Annotated[
        bool,
        typer.Option('-p', '--ppt', help=PPT_HELP),
    ] = False,
    topdf: Annotated[
        bool,
        typer.Option('--topdf', help='Convert output to PDF'),
    ] = False,
    output: Annotated[
        str | None,
        typer.Option('-o', '--output', help='Output file path'),
    ] = None,
    force: Annotated[
        bool,
        typer.Option('--force', help='Allow overwriting existing files'),
    ] = False,
    stdin: Annotated[
        bool,
        typer.Option('--stdin', help='Read document from standard input'),
    ] = False,
    stdout: Annotated[
        bool,
        typer.Option('--stdout', help='Write result to standard output'),
    ] = False,
    create: Annotated[
        bool,
        typer.Option('-cr', '--create', help='Create a blank document'),
    ] = False,
    add_text: Annotated[
        str | None,
        typer.Option('-a', '--add', help='Add text content to document'),
    ] = None,
    add_table_spec: Annotated[
        str | None,
        typer.Option(
            '--add-table',
            help='Add a table (Word). Inline "H1,H2|a1,a2" or "@file.csv"',
        ),
    ] = None,
    replace: Annotated[
        str | None,
        typer.Option('-r', '--replace', help='Replace text (use --replace-new for new value)'),
    ] = None,
    replace_new: Annotated[
        str | None,
        typer.Option('--replace-new', help='Replacement text for --replace'),
    ] = None,
    delete: Annotated[
        str | None,
        typer.Option('-d', '--delete', help='Delete content by keyword or regex'),
    ] = None,
    modify: Annotated[
        str | None,
        typer.Option('-m', '--modify', help='Modify content (use --modify-new for new value)'),
    ] = None,
    modify_new: Annotated[
        str | None,
        typer.Option('--modify-new', help='New value for --modify'),
    ] = None,
    style: Annotated[
        str | None,
        typer.Option('-s', '--style', help='Set style (key=value,key=value...)'),
    ] = None,
    template: Annotated[
        str | None,
        typer.Option('-t', '--template', help='Template data file (JSON/CSV/YAML)'),
    ] = None,
    extract: Annotated[
        str | None,
        typer.Option(
            '-x',
            '--extract',
            help='Extract data: text, tables, images, structure, or metadata',
        ),
    ] = None,
    to_md: Annotated[
        bool,
        typer.Option('--tomd', help='Convert Word to Markdown'),
    ] = False,
    to_html: Annotated[
        bool,
        typer.Option('--tohtml', help='Convert to HTML'),
    ] = False,
    to_csv: Annotated[
        bool,
        typer.Option('--to-csv', help='Export Excel to CSV'),
    ] = False,
    to_json: Annotated[
        bool,
        typer.Option('--to-json', help='Export Excel to JSON'),
    ] = False,
    heading: Annotated[
        str | None,
        typer.Option('--heading', help='Add heading (format: "level:1 text:Title")'),
    ] = None,
    latex_style: Annotated[
        bool,
        typer.Option('--latex-style', help=LATEX_STYLE_HELP),
    ] = False,
    math_formula: Annotated[
        str | None,
        typer.Option('--math', help='Add LaTeX math formula to Word document'),
    ] = None,
    regex: Annotated[
        bool,
        typer.Option('--regex', help='Enable regex for --replace and --delete'),
    ] = False,
    merge_files: Annotated[
        str | None,
        typer.Option('--merge', help='Merge multiple files (comma-separated)'),
    ] = None,
    split_mode: Annotated[
        str | None,
        typer.Option('--split', help='Split document (by-page/by-sheet/by-slide)'),
    ] = None,
    comment: Annotated[
        str | None,
        typer.Option('--comment', help='Add comment/notes (format: "cell_or_index text")'),
    ] = None,
    meta: Annotated[
        str | None,
        typer.Option('--meta', help='Set metadata (format: key=value,key=value...)'),
    ] = None,
    sheet_add: Annotated[
        str | None,
        typer.Option('--sheet-add', help='Add worksheet by name'),
    ] = None,
    sheet_delete: Annotated[
        str | None,
        typer.Option('--sheet-delete', help='Delete worksheet by name'),
    ] = None,
    sheet_rename: Annotated[
        str | None,
        typer.Option('--sheet-rename', help='Rename worksheet (format: "OldName NewName")'),
    ] = None,
    column_width: Annotated[
        str | None,
        typer.Option('--column-width', help='Set column width (format: "col=width")'),
    ] = None,
    row_height: Annotated[
        str | None,
        typer.Option('--row-height', help='Set row height (format: "row=height")'),
    ] = None,
    formula: Annotated[
        str | None,
        typer.Option('--formula', help='Set formula (format: "A1 =SUM(B1:B10)")'),
    ] = None,
    from_csv: Annotated[
        str | None,
        typer.Option('--from-csv', help='Import data from CSV file'),
    ] = None,
    sort_range: Annotated[
        str | None,
        typer.Option('--sort', help='Sort range (format: "A1:A10 asc")'),
    ] = None,
    chart_add: Annotated[
        str | None,
        typer.Option(
            '--chart-add',
            help='Add chart (format: "type=bar data=B1:C10 pos=E2")',
        ),
    ] = None,
    protect: Annotated[
        str | None,
        typer.Option('--protect', help='Set workbook password protection'),
    ] = None,
    unprotect: Annotated[
        bool,
        typer.Option('--unprotect', help='Remove workbook password protection'),
    ] = False,
    clear: Annotated[
        str | None,
        typer.Option('-cl', '--clear', help='Clear content/formats/links (Excel/Word)'),
    ] = None,
    column: Annotated[
        int,
        typer.Option('--column', help='Target column for --add (1-indexed, default 1)'),
    ] = 1,
    slide_add: Annotated[
        bool,
        typer.Option('--slide-add', help='Add a new slide'),
    ] = False,
    slide_delete: Annotated[
        int | None,
        typer.Option('--slide-delete', help='Delete slide by index (0-indexed)'),
    ] = None,
    slide_move: Annotated[
        str | None,
        typer.Option('--slide-move', help='Move slide (format: "from_index to_index")'),
    ] = None,
    slide_notes: Annotated[
        str | None,
        typer.Option('--notes', help='Add speaker notes (format: "slide_index text")'),
    ] = None,
    slide_layout: Annotated[
        str | None,
        typer.Option(
            '--layout',
            help='Apply slide layout (format: "slide_index layout_name_or_index")',
        ),
    ] = None,
    to_image: Annotated[
        bool,
        typer.Option('--toimg', help='Export PPT slides as images'),
    ] = False,
    transition: Annotated[
        str | None,
        typer.Option(
            '--transition',
            help='Set slide transition (name or "slide_index name")',
        ),
    ] = None,
    compress_media: Annotated[
        str | None,
        typer.Option(
            '--compress-media',
            help='Compress PPT images (optional "max_dimension,quality", default "1920,80")',
        ),
    ] = None,
    toc: Annotated[
        bool,
        typer.Option('--toc', help='Generate Table of Contents (Word)'),
    ] = False,
    section_break: Annotated[
        bool,
        typer.Option('--section-break', help='Insert section break (Word)'),
    ] = False,
    header_text: Annotated[
        str | None,
        typer.Option('--header', help='Set page header text (Word)'),
    ] = None,
    footer_text: Annotated[
        str | None,
        typer.Option('--footer', help='Set page footer text (Word)'),
    ] = None,
    watermark: Annotated[
        str | None,
        typer.Option('--watermark', help='Add text watermark (Word)'),
    ] = None,
    batch: Annotated[
        bool,
        typer.Option('--batch', help='Process multiple files (use --files for glob patterns)'),
    ] = False,
    files: Annotated[
        str | None,
        typer.Option(
            '--files',
            help='Glob pattern for batch processing (e.g. "reports/*.docx"). Implies --batch.',
        ),
    ] = None,
) -> None:
    """Run one-shot CLI operations on a single document or batch of files."""
    if not create and not input_file and not files and not stdin:
        console.print(app.info.help or 'tianshang-scribe — CLI Office document tool')
        return

    doc_type: DocumentType | None = None
    if word:
        doc_type = DocumentType.WORD
    elif excel:
        doc_type = DocumentType.EXCEL
    elif ppt:
        doc_type = DocumentType.PPT

    batch_files: list[str] = []
    if files:
        import glob as _glob

        batch_files = sorted(_glob.glob(files))
        if not batch_files:
            console.print(f'[red]Error:[/red] No files match pattern "{files}".')
            raise typer.Exit(code=2)
    elif input_file and any(ch in input_file for ch in '*?['):
        import glob as _glob

        batch_files = sorted(_glob.glob(input_file))
        if not batch_files:
            console.print(f'[red]Error:[/red] No files match pattern "{input_file}".')
            raise typer.Exit(code=2)
    elif input_file:
        batch_files = [input_file]

    is_batch = batch or files is not None or len(batch_files) > 1
    if is_batch and stdout:
        console.print('[red]Error:[/red] --stdout is not supported with --batch.')
        raise typer.Exit(code=2)
    if is_batch and (create or stdin):
        console.print('[red]Error:[/red] --batch is not compatible with --create/--stdin.')
        raise typer.Exit(code=2)

    def _process_one(current_input: str | None) -> None:
        resolved_type = resolve_doc_type(doc_type, current_input)

        export_ext: str | None = None
        if to_md:
            export_ext = '.md'
        elif to_html:
            export_ext = '.html'
        elif to_csv:
            export_ext = '.csv'
        elif to_json:
            export_ext = '.json'

        output_path = determine_output_path(
            current_input, output, resolved_type, topdf, to_ext=export_ext
        )

        if not check_overwrite(output_path, force):
            console.print(
                f'[red]Error:[/red] Output file "{output_path}" already exists. '
                'Use --force to overwrite.'
            )
            raise typer.Exit(code=1)

        if current_input and not (word or excel or ppt) and is_reverse_source(current_input):
            _convert_source(current_input, output_path, topdf)
            return

        try:
            if create:
                engine = create_document(resolved_type)
            elif stdin:
                import atexit
                import os
                import sys
                import tempfile

                data = sys.stdin.buffer.read()
                if not data:
                    console.print('[red]Error:[/red] No data read from stdin.')
                    raise typer.Exit(code=2)
                fd, temp_path = tempfile.mkstemp(suffix='.docx')
                os.write(fd, data)
                os.close(fd)
                atexit.register(lambda: os.unlink(temp_path))
                engine = open_document(temp_path)
            elif current_input:
                engine = open_document(current_input)
            else:
                console.print('[red]Error:[/red] Either --create or an input file is required.')
                raise typer.Exit(code=2)

            if style:
                engine.set_style(style)
                if current_input:
                    engine.apply_style_to_all()
                console.print(
                    f'[dim]Default style:[/dim] {engine.get_base_style().to_cli_string()}'
                )

            if add_text:
                if latex_style:
                    _add_latex_content(engine, add_text, column)
                else:
                    engine.add_text(add_text, column=column)

            if math_formula:
                _add_math_formula(engine, math_formula)

            if add_table_spec:
                if not hasattr(engine, 'add_table_data'):
                    console.print(
                        '[yellow]--add-table is only supported for Word documents.[/yellow]'
                    )
                else:
                    rows = parse_table_input(add_table_spec)
                    engine.add_table_data(rows)
                    console.print(f'[green]Added table[/green] {len(rows)}x{len(rows[0])}.')

            if replace and replace_new is not None:
                count = engine.replace_text(replace, replace_new, regex=regex)
                console.print(f'[green]Replaced[/green] {count} occurrence(s).')
            elif replace:
                console.print('[red]Error:[/red] --replace requires --replace-new.')
                raise typer.Exit(code=2)

            if delete:
                count = engine.replace_text(delete, '', regex=regex)
                console.print(f'[green]Deleted[/green] {count} occurrence(s).')

            if modify and modify_new is not None:
                count = engine.replace_text(modify, modify_new, regex=False)
                console.print(f'[green]Modified[/green] {count} occurrence(s).')

            if template:
                _apply_template(engine, template)

            if heading:
                _parse_heading(engine, heading)

            if meta:
                _apply_meta(engine, meta)

            if extract:
                _handle_extract(engine, extract, current_input, output)

            if merge_files and hasattr(engine, 'merge_workbooks'):
                paths = [p.strip() for p in merge_files.split(',')]
                engine.merge_workbooks(paths)
                console.print(f'[green]Merged[/green] {len(paths)} file(s).')
            elif merge_files:
                console.print('[yellow]Merge not supported for this document type.[/yellow]')

            if split_mode and hasattr(engine, 'split_by_sheet'):
                results = engine.split_by_sheet(split_mode)
                console.print(f'[green]Split[/green] {len(results)} sheet(s) to {split_mode}/.')
            elif split_mode:
                console.print('[yellow]Split not supported for this document type.[/yellow]')

            if comment and hasattr(engine, 'add_comment'):
                parts = comment.split(None, 1)
                if len(parts) == 2:
                    try:
                        engine.add_comment(int(parts[0]), parts[1])
                    except ValueError:
                        engine.add_comment(parts[0], parts[1])
                    console.print('[green]Comment added.[/green]')
                else:
                    engine.add_comment(0, comment)
                    console.print('[green]Comment added.[/green]')

            if sheet_add and hasattr(engine, 'add_sheet'):
                engine.add_sheet(sheet_add)
            if sheet_delete and hasattr(engine, 'delete_sheet'):
                engine.delete_sheet(sheet_delete)
            if sheet_rename and hasattr(engine, 'rename_sheet'):
                parts = sheet_rename.split(None, 1)
                if len(parts) == 2:
                    engine.rename_sheet(parts[0], parts[1])
            if column_width and hasattr(engine, 'set_column_width'):
                col, w = column_width.split('=', 1)
                engine.set_column_width(int(col), float(w))
            if row_height and hasattr(engine, 'set_row_height'):
                row, h = row_height.split('=', 1)
                engine.set_row_height(int(row), float(h))
            if formula and hasattr(engine, 'set_formula'):
                ref, expr = formula.split(' ', 1)
                engine.set_formula(ref, expr.strip('"').strip("'"))
            if from_csv and hasattr(engine, 'import_csv'):
                engine.import_csv(from_csv)
            if sort_range and hasattr(engine, 'sort'):
                parts = sort_range.rsplit(' ', 1)
                engine.sort(parts[0], parts[1] if len(parts) == 2 else 'asc')
            if chart_add and hasattr(engine, 'add_chart'):
                engine.add_chart(**_parse_chart_opts(chart_add))
            if protect and hasattr(engine, 'set_protection'):
                engine.set_protection(protect)
            if unprotect and hasattr(engine, 'unprotect'):
                engine.unprotect()
            if clear and hasattr(engine, 'clear_content'):
                mode = clear.lower()
                if mode == 'formats':
                    if hasattr(engine, 'clear_formats'):
                        engine.clear_formats()
                    else:
                        console.print(
                            '[yellow]Format clearing not supported for this document type.[/yellow]'
                        )
                elif mode == 'links':
                    if hasattr(engine, 'clear_links'):
                        engine.clear_links()
                    else:
                        console.print(
                            '[yellow]Link clearing not supported for this document type.[/yellow]'
                        )
                else:
                    engine.clear_content()
                console.print(f'[green]Cleared: {mode}[/green]')

            if slide_add and hasattr(engine, 'add_slide'):
                engine.add_slide()
                console.print('[green]Slide added.[/green]')
            if slide_delete is not None and hasattr(engine, 'delete_slide'):
                engine.delete_slide(slide_delete)
                console.print(f'[green]Slide {slide_delete} deleted.[/green]')
            if slide_move and hasattr(engine, 'move_slide'):
                parts = slide_move.split()
                if len(parts) == 2:
                    engine.move_slide(int(parts[0]), int(parts[1]))
                    console.print(f'[green]Slide moved from {parts[0]} to {parts[1]}.[/green]')
            if slide_notes and hasattr(engine, 'add_notes'):
                parts = slide_notes.split(None, 1)
                if len(parts) == 2:
                    engine.add_notes(int(parts[0]), parts[1])
                    console.print(f'[green]Notes added to slide {parts[0]}.[/green]')

            if slide_layout and hasattr(engine, 'apply_layout'):
                parts = slide_layout.split(None, 1)
                if len(parts) == 2:
                    engine.apply_layout(int(parts[0]), parts[1])
                    console.print(
                        f'[green]Layout "{parts[1]}" applied to slide {parts[0]}.[/green]'
                    )

            if transition and hasattr(engine, 'set_transition'):
                parts = transition.split()
                if len(parts) == 2 and parts[0].isdigit():
                    engine.set_transition(parts[1], slide_index=int(parts[0]))
                else:
                    engine.set_transition(transition)
                console.print(f'[green]Transition "{transition}" applied.[/green]')

            if compress_media and hasattr(engine, 'compress_media'):
                dims = compress_media.split(',')
                max_dimension = int(dims[0].strip()) if dims and dims[0].strip() else 1920
                quality = int(dims[1].strip()) if len(dims) > 1 and dims[1].strip() else 80
                saved = engine.compress_media(max_dimension=max_dimension, quality=quality)
                console.print(f'[green]Media compressed:[/green] saved {saved} byte(s).')

            if toc and hasattr(engine, 'add_toc'):
                engine.add_toc()
                console.print('[green]TOC generated.[/green]')
            if section_break and hasattr(engine, 'add_section_break'):
                engine.add_section_break()
                console.print('[green]Section break added.[/green]')
            if header_text and hasattr(engine, 'set_header'):
                engine.set_header(header_text)
                console.print('[green]Header set.[/green]')
            if footer_text and hasattr(engine, 'set_footer'):
                engine.set_footer(footer_text)
                console.print('[green]Footer set.[/green]')
            if watermark and hasattr(engine, 'add_watermark'):
                engine.add_watermark(watermark)
                console.print(f'[green]Watermark "{watermark}" added.[/green]')

            if to_image and hasattr(engine, 'to_images'):
                results = engine.to_images(output)
                console.print(f'[green]{len(results)} slide images saved to {output}/.[/green]')
            elif topdf:
                engine.to_pdf(output_path)
                console.print(f'[green]PDF saved:[/green] {output_path}')
            elif to_md:
                import tempfile

                from src.transform.pdf import word_to_markdown

                with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tf:
                    engine.save(tf.name)
                    try:
                        word_to_markdown(tf.name, str(output_path))
                    except Exception as e:
                        console.print(f'[yellow]Markdown conversion skipped:[/yellow] {e}')
                        engine.save(output_path)
                import os

                os.unlink(tf.name)
                console.print(f'[green]Markdown saved:[/green] {output_path}')
            elif to_csv and hasattr(engine, 'export_csv'):
                engine.export_csv(output_path)
                console.print(f'[green]CSV saved:[/green] {output_path}')
            elif to_json and hasattr(engine, 'export_json'):
                engine.export_json(output_path)
                console.print(f'[green]JSON saved:[/green] {output_path}')
            elif to_html:
                if hasattr(engine, 'export_html'):
                    engine.export_html(output_path)
                    console.print(f'[green]HTML saved:[/green] {output_path}')
                else:
                    import tempfile

                    from src.transform.pdf import word_to_html

                    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tf:
                        engine.save(tf.name)
                        try:
                            word_to_html(tf.name, str(output_path))
                        except Exception as e:
                            console.print(f'[yellow]HTML conversion skipped:[/yellow] {e}')
                            engine.save(output_path)
                    import os

                    os.unlink(tf.name)
                    console.print(f'[green]HTML saved:[/green] {output_path}')
            elif stdout:
                import atexit
                import os
                import sys
                import tempfile

                fd, temp_path = tempfile.mkstemp(suffix='.docx')
                os.close(fd)
                atexit.register(lambda: os.unlink(temp_path))
                engine.save(temp_path)
                with open(temp_path, 'rb') as f:
                    sys.stdout.buffer.write(f.read())
                console.print('[green]Output written to stdout.[/green]')
            else:
                engine.save(output_path)
                console.print(f'[green]Saved:[/green] {output_path}')

        except typer.Exit:
            raise
        except FileNotFoundError as e:
            console.print(f'[red]Error:[/red] {e}')
            raise typer.Exit(code=1) from None
        except ValueError as e:
            console.print(f'[red]Error:[/red] {e}')
            raise typer.Exit(code=2) from None
        except NotImplementedError as e:
            console.print(f'[yellow]Not implemented:[/yellow] {e}')
            raise typer.Exit(code=3) from None
        except Exception as e:
            console.print(f'[red]Unexpected error:[/red] {e}')
            raise typer.Exit(code=1) from None

    if not is_batch:
        _process_one(input_file)
        return

    succeeded = 0
    failed: list[str] = []
    for current_input in batch_files:
        try:
            _process_one(current_input)
            succeeded += 1
        except typer.Exit:
            failed.append(current_input)
    console.print(f'[bold]Batch complete:[/bold] {succeeded} succeeded, {len(failed)} failed.')
    if failed:
        for f in failed:
            console.print(f'  [red]x[/red] {f}')
        raise typer.Exit(code=1)


def _convert_source(input_path: str, output_path: str, to_pdf: bool = False) -> None:
    suffix = Path(input_path).suffix.lower()
    intermediate = output_path
    if to_pdf:
        intermediate = str(Path(output_path).with_suffix('.docx'))

    if suffix == '.json':
        from src.core.excel_engine import ExcelEngine

        engine = ExcelEngine()
        engine.create()
        engine.import_json(input_path)
        engine.save(intermediate)
        console.print(f'[green]Imported JSON:[/green] {input_path} -> {intermediate}')
    elif suffix in ('.md', '.markdown', '.html', '.htm'):
        from src.transform.reverse import html_to_word, markdown_to_word

        if suffix in ('.md', '.markdown'):
            markdown_to_word(input_path, intermediate)
        else:
            html_to_word(input_path, intermediate)
        console.print(f'[green]Converted:[/green] {input_path} -> {intermediate}')
    else:
        raise ValueError(f'Unsupported source format: {suffix}')

    if to_pdf:
        from src.transform.pdf import word_to_pdf

        word_to_pdf(intermediate, output_path)
        console.print(f'[green]Converted to PDF:[/green] {output_path}')


def _add_latex_content(engine: Any, text: str, column: int = 1) -> None:
    if hasattr(engine, 'add_latex_content'):
        engine.add_latex_content(text)
        console.print('[green]LaTeX content added with style parsing.[/green]')
    else:
        console.print(
            '[yellow]Warning:[/yellow] LaTeX style parsing not supported for this document type.'
        )
        engine.add_text(text, column=column)


def _add_math_formula(engine: Any, latex: str) -> None:
    if hasattr(engine, 'add_math_formula'):
        engine.add_math_formula(latex)
        console.print(f'[green]Math formula added:[/green] ${latex}$')
    else:
        console.print('[red]Error:[/red] --math is only supported for Word documents.')
        raise typer.Exit(code=2)


def _apply_template(engine: Any, template_path: str) -> None:
    tpl_engine = TemplateEngine(template_path)
    count = tpl_engine.fill(engine)
    console.print(f'[green]Template filled:[/green] {count} placeholder(s) replaced.')


def _parse_heading(engine: Any, heading_str: str) -> None:
    import re

    match = re.match(r'level:(\d+)\s+text:(.+)', heading_str)
    if match:
        level = int(match.group(1))
        text = match.group(2)
        if hasattr(engine, 'add_heading'):
            engine.add_heading(text, level=level)
        else:
            engine.add_text(text)
    else:
        engine.add_text(heading_str)


def _apply_meta(engine: Any, meta_str: str) -> None:
    kwargs: dict[str, str] = {}
    pairs = [p.strip() for p in meta_str.split(',') if p.strip()]
    for pair in pairs:
        if '=' in pair:
            key, _, value = pair.partition('=')
            kwargs[key.strip()] = value.strip()
    engine.set_metadata(**kwargs)


def _parse_chart_opts(opts_str: str) -> dict[str, str]:
    result: dict[str, str] = {}
    for pair in opts_str.replace(' ', ',').split(','):
        if '=' in pair:
            k, v = pair.split('=', 1)
            result[k] = v
    return result


def _handle_extract(
    engine: Any,
    mode: str,
    input_path: str | None = None,
    output_dir: str | None = None,
) -> None:
    mode = (mode or '').lower()
    if mode == 'metadata':
        import json

        meta = engine.get_metadata()
        console.print_json(json.dumps(meta, ensure_ascii=False, default=str))
    elif mode == 'text':
        console.print(engine.extract_text())
    elif mode == 'structure':
        import json

        console.print_json(json.dumps(engine.extract_structure(), ensure_ascii=False, default=str))
    elif mode == 'tables':
        import csv
        import io

        tables = engine.extract_tables()
        if not tables:
            console.print('[yellow]No tables found.[/yellow]')
            return
        buf = io.StringIO()
        writer = csv.writer(buf)
        for table in tables:
            for row in table:
                writer.writerow(row)
            writer.writerow([])
        console.print(buf.getvalue())
    elif mode == 'images':
        if not output_dir:
            if input_path:
                base = Path(input_path).stem
                output_dir = str(Path(input_path).with_name(f'{base}-images'))
            else:
                output_dir = 'images'
        saved = engine.extract_images(output_dir)
        console.print(f'[green]Extracted[/green] {len(saved)} image(s) to {output_dir}/.')
        for s in saved:
            console.print(f'  {s}')
    else:
        console.print(
            f'[yellow]Extract mode "{mode}" not supported. '
            'Use text, tables, images, structure, or metadata.[/yellow]'
        )


open_app = typer.Typer(
    name='open',
    help='Open a document in an interactive editing session',
    add_completion=False,
)


def _open_engine(path: str, explicit: DocumentType | None) -> Any:
    """Open ``path``, optionally forcing the document type."""
    if explicit is not None and detect_document_type(path) != explicit:
        from src.core.excel_engine import ExcelEngine
        from src.core.ppt_engine import PptEngine
        from src.core.word_engine import WordEngine

        engine_map: dict[DocumentType, type[WordEngine] | type[ExcelEngine] | type[PptEngine]] = {
            DocumentType.WORD: WordEngine,
            DocumentType.EXCEL: ExcelEngine,
            DocumentType.PPT: PptEngine,
        }
        engine_cls = engine_map.get(explicit)
        if engine_cls is None:
            raise ValueError(f'Unsupported document type: {explicit}')
        engine = engine_cls()
        engine.open(path)
        return engine
    return open_document(path)


@open_app.command()
def open_cmd(
    file: Annotated[str, typer.Argument(help='Document path to open')],
    latex_style: Annotated[
        bool,
        typer.Option('--latex-style', help=LATEX_STYLE_HELP),
    ] = False,
    word: Annotated[
        bool,
        typer.Option('-w', '--word', help=WORD_HELP),
    ] = False,
    excel: Annotated[
        bool,
        typer.Option('-e', '--excel', help=EXCEL_HELP),
    ] = False,
    ppt: Annotated[
        bool,
        typer.Option('-p', '--ppt', help=PPT_HELP),
    ] = False,
) -> None:
    """Open a document and enter an interactive editing session."""
    explicit: DocumentType | None = None
    if word:
        explicit = DocumentType.WORD
    elif excel:
        explicit = DocumentType.EXCEL
    elif ppt:
        explicit = DocumentType.PPT
    engine = _open_engine(file, explicit)
    from src.cli.repl import InteractiveSession

    InteractiveSession(engine, file, console, latex_style=latex_style).run()


def main_cli() -> None:
    """Entry point dispatching between the ``open`` subcommand and the main app."""
    import sys

    args = sys.argv[1:]
    if args and args[0].lower() == 'open':
        open_app(sys.argv[2:])
    else:
        app()


if __name__ == '__main__':
    main_cli()
