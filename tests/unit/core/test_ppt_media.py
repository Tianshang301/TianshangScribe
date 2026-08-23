"""Tests for 0.9.0 P1 media insertion (add_movie / add_audio)."""

from __future__ import annotations

from pathlib import Path

import pytest

from tianshang_scribe.core.ppt_engine import PptEngine


@pytest.fixture
def engine() -> PptEngine:
    e = PptEngine()
    e.create()
    return e


def _media_file(tmp_path: Path, name: str, magic: bytes) -> Path:
    p = tmp_path / name
    p.write_bytes(magic + b'\x00' * 64)
    return p


class TestAddMedia:
    def test_add_movie_mp4(self, engine: PptEngine, tmp_path: Path) -> None:
        engine.add_slide()
        mp4 = _media_file(tmp_path, 'clip.mp4', b'\x00\x00\x00 ftypisom')
        shape = engine.add_movie(0, mp4, left=0.5, top=0.5)
        assert 'MEDIA' in str(shape.shape_type)
        assert shape.left == pytest.approx(0.5 * 914400)
        out = tmp_path / 'deck.pptx'
        engine.save(out)
        reopened = PptEngine()
        reopened.open(out)
        assert any('MEDIA' in str(s.shape_type) for s in reopened.prs.slides[0].shapes)

    def test_add_movie_unknown_ext_falls_back(self, engine: PptEngine, tmp_path: Path) -> None:
        engine.add_slide()
        clip = _media_file(tmp_path, 'clip.weird', b'MEDIA')
        shape = engine.add_movie(0, clip)  # mime falls back to video/unknown
        assert 'MEDIA' in str(shape.shape_type)

    def test_add_movie_poster_frame(self, engine: PptEngine, tmp_path: Path) -> None:
        from PIL import Image

        engine.add_slide()
        mp4 = _media_file(tmp_path, 'clip.mp4', b'\x00\x00\x00 ftypisom')
        poster = tmp_path / 'poster.png'
        Image.new('RGB', (8, 8), color=(10, 20, 30)).save(poster)
        engine.add_movie(0, mp4, poster=poster)
        # poster becomes an additional picture part in the package
        image_parts = [
            part
            for part in engine.prs.part.package.iter_parts()
            if part.content_type == 'image/png'
        ]
        assert image_parts

    def test_add_movie_missing_file_raises(self, engine: PptEngine, tmp_path: Path) -> None:
        engine.add_slide()
        with pytest.raises(FileNotFoundError, match='media file not found'):
            engine.add_movie(0, tmp_path / 'ghost.mp4')

    def test_add_movie_bad_slide_index(self, engine: PptEngine, tmp_path: Path) -> None:
        engine.add_slide()
        mp4 = _media_file(tmp_path, 'clip.mp4', b'\x00\x00\x00 ftyp')
        with pytest.raises(IndexError, match='out of range'):
            engine.add_movie(7, mp4)

    def test_add_audio_mp3(self, engine: PptEngine, tmp_path: Path) -> None:
        engine.add_slide()
        mp3 = _media_file(tmp_path, 'sound.mp3', b'ID3')
        shape = engine.add_audio(0, mp3)
        assert 'MEDIA' in str(shape.shape_type)
        audio_parts = [
            part
            for part in engine.prs.part.package.iter_parts()
            if getattr(part, 'content_type', '') == 'audio/mpeg'
        ]
        assert audio_parts

    @pytest.mark.parametrize('bad', ['song.flac', 'voice.ogg'])
    def test_add_audio_unsupported_format(
        self, engine: PptEngine, tmp_path: Path, bad: str
    ) -> None:
        engine.add_slide()
        f = _media_file(tmp_path, bad, b'data')
        with pytest.raises(ValueError, match='Unsupported audio format'):
            engine.add_audio(0, f)
