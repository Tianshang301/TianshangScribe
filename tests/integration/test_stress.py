"""Stress / benchmark tests (CI ``-m stress``).

These run only in the dedicated ``stress`` CI job, not in the default suite
(which deselects them via ``-m "not stress"``).
"""

from __future__ import annotations

import os
import tempfile

import pytest


@pytest.mark.stress
@pytest.mark.benchmark(group='office-roundtrip', min_rounds=5)
def test_office_document_roundtrip_benchmark(benchmark) -> None:
    """Benchmark create -> add -> save for Word, Excel, and PowerPoint."""

    def _roundtrip() -> None:
        from tianshang_scribe.core.excel_engine import ExcelEngine
        from tianshang_scribe.core.ppt_engine import PptEngine
        from tianshang_scribe.core.word_engine import WordEngine

        for suffix, engine_cls in (
            ('.docx', WordEngine),
            ('.xlsx', ExcelEngine),
            ('.pptx', PptEngine),
        ):
            fd, path = tempfile.mkstemp(suffix=suffix)
            os.close(fd)
            try:
                engine = engine_cls()
                engine.create()
                engine.add_text('stress benchmark content ' * 20)
                engine.save(path)
            finally:
                os.unlink(path)

    benchmark(_roundtrip)
