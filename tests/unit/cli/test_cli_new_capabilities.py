"""Baseline tests for the new Excel/PPT CLI capability options.

Exercises the one-shot CLI end-to-end with real engines and verifies the
resulting document state (no mocks).
"""

from __future__ import annotations

from pathlib import Path

from typer.testing import CliRunner

from tianshang_scribe.cli.main import app
from tianshang_scribe.core.document import open_document

runner = CliRunner()


def test_cli_excel_freeze_and_number_format(tmp_path: Path) -> None:
    out = tmp_path / 'book.xlsx'
    result = runner.invoke(
        app,
        ['-e', '-cr', '-o', str(out), '--freeze', 'A2', '--number-format', 'A1:A2=0.00%'],
    )
    assert result.exit_code == 0, result.output
    eng = open_document(out)
    assert eng.wb.active.freeze_panes == 'A2'
    assert eng.wb.active['A1'].number_format == '0.00%'


def test_cli_excel_conditional_format_and_data_validation(tmp_path: Path) -> None:
    out = tmp_path / 'book2.xlsx'
    result = runner.invoke(
        app,
        [
            '-e',
            '-cr',
            '-o',
            str(out),
            '--conditional-format',
            'B1:B5=color_scale',
            '--data-validation',
            'C1:C5=list:yes,no',
        ],
    )
    assert result.exit_code == 0, result.output
    eng = open_document(out)
    assert len(list(eng.wb.active.conditional_formatting)) >= 1
    types = {dv.type for dv in eng.wb.active.data_validations.dataValidation}
    assert 'list' in types


def test_cli_ppt_table_and_chart(tmp_path: Path) -> None:
    out = tmp_path / 'deck.pptx'
    result = runner.invoke(
        app,
        [
            '-p',
            '-cr',
            '-o',
            str(out),
            '--ppt-table',
            'H1,H2|a1,a2',
            '--ppt-chart',
            'bar|S1,S2|Cat1,1,2|Cat2,3,4',
        ],
    )
    assert result.exit_code == 0, result.output
    eng = open_document(out)
    shapes = list(eng.prs.slides[0].shapes)
    assert any(sh.has_table for sh in shapes)
    assert any(sh.has_chart for sh in shapes)
