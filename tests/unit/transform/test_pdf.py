"""Unit tests for PDF / Markdown / HTML conversion (tianshang_scribe/transform/pdf.py)."""

from __future__ import annotations

from pathlib import Path

import pytest

from tianshang_scribe.transform import pdf


@pytest.fixture
def input_file(tmp_path: Path) -> Path:
    p = tmp_path / 'input.docx'
    p.write_bytes(b'x' * 100)
    return p


def _mock_which(monkeypatch, path: str) -> None:
    monkeypatch.setattr(pdf.shutil, 'which', lambda name: path if name == 'office2pdf' else None)


class TestEngineDiscovery:
    def test_find_office2pdf(self, monkeypatch) -> None:
        monkeypatch.setattr(
            pdf.shutil,
            'which',
            lambda name: '/usr/bin/office2pdf' if name == 'office2pdf' else None,
        )
        assert pdf._find_office2pdf() == '/usr/bin/office2pdf'

    def test_find_office2pdf_missing(self, monkeypatch) -> None:
        monkeypatch.setattr(pdf.shutil, 'which', lambda name: None)
        assert pdf._find_office2pdf() is None

    def test_find_libreoffice_via_which(self, monkeypatch) -> None:
        def which(name: str) -> str | None:
            return '/usr/bin/libreoffice' if name == 'libreoffice' else None

        monkeypatch.setattr(pdf.shutil, 'which', which)
        assert pdf._find_libreoffice() == '/usr/bin/libreoffice'

    def test_find_libreoffice_soffice_alias(self, monkeypatch) -> None:
        def which(name: str) -> str | None:
            return '/usr/bin/soffice' if name == 'soffice' else None

        monkeypatch.setattr(pdf.shutil, 'which', which)
        assert pdf._find_libreoffice() == '/usr/bin/soffice'

    def test_find_libreoffice_windows_path(self, monkeypatch, tmp_path: Path) -> None:
        windows_path = r'C:\Program Files\LibreOffice\program\soffice.exe'
        monkeypatch.setattr(pdf.shutil, 'which', lambda name: None)
        monkeypatch.setattr(Path, 'exists', lambda self: str(self) == windows_path)
        assert pdf._find_libreoffice() == windows_path

    def test_find_libreoffice_missing(self, monkeypatch) -> None:
        monkeypatch.setattr(pdf.shutil, 'which', lambda name: None)
        monkeypatch.setattr(Path, 'exists', lambda self: False)
        assert pdf._find_libreoffice() is None


class TestOffice2pdf:
    def test_convert_runs_subprocess(self, monkeypatch, input_file: Path, tmp_path: Path) -> None:
        out = tmp_path / 'out.pdf'
        calls: list[list[str]] = []
        _mock_which(monkeypatch, '/usr/bin/office2pdf')

        def fake_run(cmd: list[str], **kwargs: object) -> None:
            calls.append(cmd)
            out.write_bytes(b'pdf')

        monkeypatch.setattr(pdf.subprocess, 'run', fake_run)
        pdf._convert_via_office2pdf(input_file, out)
        assert calls == [['/usr/bin/office2pdf', str(input_file), '-o', str(out)]]

    def test_missing_binary_raises(self, monkeypatch, input_file: Path, tmp_path: Path) -> None:
        monkeypatch.setattr(pdf.shutil, 'which', lambda name: None)
        with pytest.raises(RuntimeError, match='office2pdf not found'):
            pdf._convert_via_office2pdf(input_file, tmp_path / 'out.pdf')


class TestLibreOffice:
    def test_convert_runs_headless(self, monkeypatch, input_file: Path, tmp_path: Path) -> None:
        out = tmp_path / 'out.pdf'
        calls: list[list[str]] = []
        monkeypatch.setattr(
            pdf.shutil,
            'which',
            lambda name: '/usr/bin/libreoffice' if name == 'libreoffice' else None,
        )

        def fake_run(cmd: list[str], **kwargs: object) -> None:
            calls.append(cmd)

        monkeypatch.setattr(pdf.subprocess, 'run', fake_run)
        pdf._convert_via_libreoffice(input_file, out)
        assert calls == [
            [
                '/usr/bin/libreoffice',
                '--headless',
                '--convert-to',
                'pdf',
                '--outdir',
                str(tmp_path),
                str(input_file),
            ]
        ]

    def test_missing_binary_raises(self, monkeypatch, input_file: Path, tmp_path: Path) -> None:
        monkeypatch.setattr(pdf.shutil, 'which', lambda name: None)
        monkeypatch.setattr(Path, 'exists', lambda self: False)
        with pytest.raises(RuntimeError, match='No PDF engine found'):
            pdf._convert_via_libreoffice(input_file, tmp_path / 'out.pdf')


class TestConvert:
    def test_no_engine_raises(self, monkeypatch, input_file: Path, tmp_path: Path) -> None:
        monkeypatch.setattr(pdf.shutil, 'which', lambda name: None)
        monkeypatch.setattr(Path, 'exists', lambda self: False)
        with pytest.raises(RuntimeError, match='No PDF conversion engine found'):
            pdf._convert(input_file, tmp_path / 'out.pdf')

    def test_office2pdf_selected(self, monkeypatch, input_file: Path, tmp_path: Path) -> None:
        out = tmp_path / 'out.pdf'
        _mock_which(monkeypatch, '/usr/bin/office2pdf')
        calls: list[list[str]] = []

        def fake_run(cmd: list[str], **kwargs: object) -> None:
            calls.append(cmd)
            out.write_bytes(b'pdf')

        monkeypatch.setattr(pdf.subprocess, 'run', fake_run)
        pdf._convert(input_file, out)
        assert calls[0][0] == '/usr/bin/office2pdf'

    def test_libreoffice_fallback(self, monkeypatch, input_file: Path, tmp_path: Path) -> None:
        out = tmp_path / 'out.pdf'
        calls: list[list[str]] = []

        def which(name: str) -> str | None:
            return '/usr/bin/soffice' if name == 'soffice' else None

        monkeypatch.setattr(pdf.shutil, 'which', which)

        def fake_run(cmd: list[str], **kwargs: object) -> None:
            calls.append(cmd)
            out.write_bytes(b'pdf')

        monkeypatch.setattr(pdf.subprocess, 'run', fake_run)
        pdf._convert(input_file, out)
        assert calls[0][0] == '/usr/bin/soffice'


class TestPublicConverters:
    def test_word_to_pdf(self, monkeypatch, input_file: Path, tmp_path: Path) -> None:
        out = tmp_path / 'out.pdf'
        _mock_which(monkeypatch, '/usr/bin/office2pdf')

        def fake_run(cmd: list[str], **kwargs: object) -> None:
            out.write_bytes(b'pdf')

        monkeypatch.setattr(pdf.subprocess, 'run', fake_run)
        pdf.word_to_pdf(input_file, out)
        assert out.exists()

    def test_excel_to_pdf(self, monkeypatch, input_file: Path, tmp_path: Path) -> None:
        out = tmp_path / 'o.pdf'
        monkeypatch.setattr(
            pdf.shutil, 'which', lambda name: '/usr/bin/soffice' if name == 'soffice' else None
        )
        monkeypatch.setattr(pdf.subprocess, 'run', lambda cmd, **kw: out.write_bytes(b'pdf'))
        pdf.excel_to_pdf(input_file, out)
        assert out.exists()

    def test_ppt_to_pdf(self, monkeypatch, input_file: Path, tmp_path: Path) -> None:
        out = tmp_path / 'o.pdf'
        monkeypatch.setattr(
            pdf.shutil, 'which', lambda name: '/usr/bin/soffice' if name == 'soffice' else None
        )
        monkeypatch.setattr(pdf.subprocess, 'run', lambda cmd, **kw: out.write_bytes(b'pdf'))
        pdf.ppt_to_pdf(input_file, out)
        assert out.exists()


def _disable_mammoth(monkeypatch) -> None:
    import builtins

    real_import = builtins.__import__

    def fake_import(name: str, *a: object, **k: object) -> object:
        if name == 'mammoth' or name.startswith('mammoth.'):
            raise ImportError(name)
        return real_import(name, *a, **k)

    monkeypatch.setattr(builtins, '__import__', fake_import)


class TestMarkdownHtml:
    def test_word_to_markdown_via_mammoth(
        self, monkeypatch, input_file: Path, tmp_path: Path
    ) -> None:
        import mammoth

        out = tmp_path / 'out.md'
        monkeypatch.setattr(
            mammoth,
            'convert_to_markdown',
            lambda f: type('R', (), {'value': '# Hi'})(),
        )
        monkeypatch.setattr(pdf, 'open', open, raising=False)
        pdf.word_to_markdown(input_file, out)
        assert out.read_text() == '# Hi'

    def test_word_to_markdown_missing_deps(
        self, monkeypatch, input_file: Path, tmp_path: Path
    ) -> None:
        _disable_mammoth(monkeypatch)
        monkeypatch.setattr(pdf.shutil, 'which', lambda name: None)
        with pytest.raises(RuntimeError, match='pandoc not found'):
            pdf.word_to_markdown(input_file, tmp_path / 'out.md')

    def test_word_to_markdown_pandoc_fallback(
        self, monkeypatch, input_file: Path, tmp_path: Path
    ) -> None:
        out = tmp_path / 'out.md'
        calls: list[list[str]] = []
        _disable_mammoth(monkeypatch)
        monkeypatch.setattr(pdf.shutil, 'which', lambda name: '/usr/bin/pandoc')

        def fake_run(cmd: list[str], **kwargs: object) -> None:
            calls.append(cmd)
            out.write_text('md')

        monkeypatch.setattr(pdf.subprocess, 'run', fake_run)
        pdf.word_to_markdown(input_file, out)
        assert calls[0][0] == '/usr/bin/pandoc'
        assert out.read_text() == 'md'

    def test_word_to_html_via_mammoth(self, monkeypatch, input_file: Path, tmp_path: Path) -> None:
        import mammoth

        out = tmp_path / 'out.html'
        monkeypatch.setattr(
            mammoth, 'convert_to_html', lambda f: type('R', (), {'value': '<p>Hi</p>'})()
        )
        pdf.word_to_html(input_file, out)
        assert out.read_text() == '<p>Hi</p>'

    def test_word_to_html_pandoc_fallback(
        self, monkeypatch, input_file: Path, tmp_path: Path
    ) -> None:
        out = tmp_path / 'out.html'
        _disable_mammoth(monkeypatch)
        monkeypatch.setattr(pdf.shutil, 'which', lambda name: '/usr/bin/pandoc')
        monkeypatch.setattr(pdf.subprocess, 'run', lambda cmd, **kw: out.write_text('html'))
        pdf.word_to_html(input_file, out)
        assert out.read_text() == 'html'
