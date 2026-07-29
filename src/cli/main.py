from __future__ import annotations

from typing import Annotated, Optional

import typer
from rich.console import Console

from src.cli.global_opts import (
    check_overwrite,
    determine_output_path,
    resolve_doc_type,
)
from src.core.document import DocumentType, create_document, open_document
from src.rendering.template import TemplateEngine

app = typer.Typer(
    name='tianshang-scribe',
    help='天殇·书契 — Cross-platform CLI Office document processing tool',
)
console = Console()


@app.callback(invoke_without_command=True)
def main(
    input_file: Annotated[
        Optional[str],
        typer.Argument(help='Input document path (omit with --create)'),
    ] = None,
    word: Annotated[
        bool,
        typer.Option('-w', '--word', help='Process Word document'),
    ] = False,
    excel: Annotated[
        bool,
        typer.Option('-e', '--excel', help='Process Excel workbook'),
    ] = False,
    ppt: Annotated[
        bool,
        typer.Option('-p', '--ppt', help='Process PowerPoint presentation'),
    ] = False,
    topdf: Annotated[
        bool,
        typer.Option('--topdf', help='Convert output to PDF'),
    ] = False,
    output: Annotated[
        Optional[str],
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
        Optional[str],
        typer.Option('-a', '--add', help='Add text content to document'),
    ] = None,
    replace: Annotated[
        Optional[str],
        typer.Option('-r', '--replace', help='Replace text (use --replace-new for new value)'),
    ] = None,
    replace_new: Annotated[
        Optional[str],
        typer.Option('--replace-new', help='Replacement text for --replace'),
    ] = None,
    delete: Annotated[
        Optional[str],
        typer.Option('-d', '--delete', help='Delete content by keyword or regex'),
    ] = None,
    modify: Annotated[
        Optional[str],
        typer.Option('-m', '--modify', help='Modify content (use --modify-new for new value)'),
    ] = None,
    modify_new: Annotated[
        Optional[str],
        typer.Option('--modify-new', help='New value for --modify'),
    ] = None,
    style: Annotated[
        Optional[str],
        typer.Option('-s', '--style', help='Set style (key=value,key=value...)'),
    ] = None,
    template: Annotated[
        Optional[str],
        typer.Option('-t', '--template', help='Template data file (JSON/CSV/YAML)'),
    ] = None,
    extract: Annotated[
        Optional[str],
        typer.Option('-x', '--extract', help='Extract data (text, metadata)'),
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
        Optional[str],
        typer.Option('--heading', help='Add heading (format: "level:1 text:Title")'),
    ] = None,
    latex_style: Annotated[
        bool,
        typer.Option('--latex-style', help='Enable LaTeX style markup parsing'),
    ] = False,
    math_formula: Annotated[
        Optional[str],
        typer.Option('--math', help='Add LaTeX math formula to Word document'),
    ] = None,
    regex: Annotated[
        bool,
        typer.Option('--regex', help='Enable regex for --replace and --delete'),
    ] = False,
    merge_files: Annotated[
        Optional[str],
        typer.Option('--merge', help='Merge multiple files (comma-separated)'),
    ] = None,
    meta: Annotated[
        Optional[str],
        typer.Option('--meta', help='Set metadata (format: key=value,key=value...)'),
    ] = None,
    sheet_add: Annotated[
        Optional[str],
        typer.Option('--sheet-add', help='Add worksheet by name'),
    ] = None,
    sheet_delete: Annotated[
        Optional[str],
        typer.Option('--sheet-delete', help='Delete worksheet by name'),
    ] = None,
    sheet_rename: Annotated[
        Optional[str],
        typer.Option('--sheet-rename', help='Rename worksheet (format: "OldName NewName")'),
    ] = None,
    column_width: Annotated[
        Optional[str],
        typer.Option('--column-width', help='Set column width (format: "col=width")'),
    ] = None,
    row_height: Annotated[
        Optional[str],
        typer.Option('--row-height', help='Set row height (format: "row=height")'),
    ] = None,
    formula: Annotated[
        Optional[str],
        typer.Option('--formula', help='Set formula (format: "A1 =SUM(B1:B10)")'),
    ] = None,
    from_csv: Annotated[
        Optional[str],
        typer.Option('--from-csv', help='Import data from CSV file'),
    ] = None,
    sort_range: Annotated[
        Optional[str],
        typer.Option('--sort', help='Sort range (format: "A1:A10 asc")'),
    ] = None,
    chart_add: Annotated[
        Optional[str],
        typer.Option(
            '--chart-add',
            help='Add chart (format: "type=bar data=B1:C10 pos=E2")',
        ),
    ] = None,
    protect: Annotated[
        Optional[str],
        typer.Option('--protect', help='Set workbook password protection'),
    ] = None,
    unprotect: Annotated[
        bool,
        typer.Option('--unprotect', help='Remove workbook password protection'),
    ] = False,
    clear: Annotated[
        bool,
        typer.Option('--clear', help='Clear all cell content'),
    ] = False,
    column: Annotated[
        int,
        typer.Option('--column', help='Target column for --add (1-indexed, default 1)'),
    ] = 1,
) -> None:
    if not create and not input_file and not stdin:
        console.print(app.get_help())
        return

    doc_type: DocumentType | None = None
    if word:
        doc_type = DocumentType.WORD
    elif excel:
        doc_type = DocumentType.EXCEL
    elif ppt:
        doc_type = DocumentType.PPT

    resolved_type = resolve_doc_type(doc_type, input_file)

    output_path = determine_output_path(
        input_file, output, resolved_type, topdf
    )

    if not check_overwrite(output_path, force):
        console.print(
            f'[red]Error:[/red] Output file "{output_path}" already exists. '
            'Use --force to overwrite.'
        )
        raise typer.Exit(code=1)

    try:
        if create:
            engine = create_document(resolved_type)
        elif input_file:
            engine = open_document(input_file)
        else:
            console.print('[red]Error:[/red] Either --create or an input file is required.')
            raise typer.Exit(code=2)

        if style:
            engine.set_style(style)
            if input_file:
                engine.apply_style_to_all()
            console.print(f'[dim]Default style:[/dim] {engine.get_base_style().to_cli_string()}')

        if add_text:
            if latex_style:
                _add_latex_content(engine, add_text, column)
            else:
                engine.add_text(add_text, column=column)

        if math_formula:
            _add_math_formula(engine, math_formula)

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
            engine.clear_content()

        if topdf:
            engine.to_pdf(output_path)
            console.print(f'[green]PDF saved:[/green] {output_path}')
        elif to_csv and hasattr(engine, 'export_csv'):
            engine.export_csv(output_path)
            console.print(f'[green]CSV saved:[/green] {output_path}')
        elif to_json and hasattr(engine, 'export_json'):
            engine.export_json(output_path)
            console.print(f'[green]JSON saved:[/green] {output_path}')
        elif to_html and hasattr(engine, 'export_html'):
            engine.export_html(output_path)
            console.print(f'[green]HTML saved:[/green] {output_path}')
        else:
            engine.save(output_path)
            console.print(f'[green]Saved:[/green] {output_path}')

    except FileNotFoundError as e:
        console.print(f'[red]Error:[/red] {e}')
        raise typer.Exit(code=1)
    except ValueError as e:
        console.print(f'[red]Error:[/red] {e}')
        raise typer.Exit(code=2)
    except NotImplementedError as e:
        console.print(f'[yellow]Not implemented:[/yellow] {e}')
        raise typer.Exit(code=3)
    except Exception as e:
        console.print(f'[red]Unexpected error:[/red] {e}')
        raise typer.Exit(code=1)


def _add_latex_content(engine, text: str, column: int = 1) -> None:
    if hasattr(engine, 'add_latex_content'):
        engine.add_latex_content(text)
        console.print('[green]LaTeX content added with style parsing.[/green]')
    else:
        console.print(
            '[yellow]Warning:[/yellow] '
            'LaTeX style parsing not supported for this document type.'
        )
        engine.add_text(text, column=column)


def _add_math_formula(engine, latex: str) -> None:
    if hasattr(engine, 'add_math_formula'):
        engine.add_math_formula(latex)
        console.print(f'[green]Math formula added:[/green] ${latex}$')
    else:
        console.print('[red]Error:[/red] --math is only supported for Word documents.')
        raise typer.Exit(code=2)


def _apply_template(engine, template_path: str) -> None:
    tpl_engine = TemplateEngine(template_path)
    count = tpl_engine.fill(engine)
    console.print(f'[green]Template filled:[/green] {count} placeholder(s) replaced.')


def _parse_heading(engine, heading_str: str) -> None:
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


def _apply_meta(engine, meta_str: str) -> None:
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


def main_cli() -> None:
    app()


if __name__ == '__main__':
    main_cli()
