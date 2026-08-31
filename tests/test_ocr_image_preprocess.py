#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from io import BytesIO
from pathlib import Path
from unittest.mock import patch

from PIL import Image

from core.ocr.image_preprocess import preprocess_image_for_ocr


MIN_SHORT_SIDE = 1000


def _save_rgb(path: Path, width: int, height: int, color=(180, 180, 180)) -> None:
    Image.new("RGB", (width, height), color=color).save(path, format="PNG")


def _png_size(data: bytes) -> tuple[int, int]:
    with Image.open(BytesIO(data)) as image:
        return image.size


def test_short_side_above_threshold_is_noop(tmp_path: Path) -> None:
    path = tmp_path / "large.png"
    _save_rgb(path, width=1600, height=1200)
    before = path.read_bytes()

    result = preprocess_image_for_ocr(str(path), min_short_side=MIN_SHORT_SIDE)

    assert result is not None
    assert result.applied is False
    assert result.image_data == before
    assert result.mime_type == "image/png"
    assert path.read_bytes() == before


def test_short_side_equal_to_threshold_is_noop(tmp_path: Path) -> None:
    path = tmp_path / "equal.png"
    _save_rgb(path, width=1400, height=1000)
    before = path.read_bytes()

    result = preprocess_image_for_ocr(str(path), min_short_side=MIN_SHORT_SIDE)

    assert result is not None
    assert result.applied is False
    assert result.image_data == before
    assert path.read_bytes() == before


def test_short_side_below_threshold_returns_2x_png(tmp_path: Path) -> None:
    path = tmp_path / "small.png"
    _save_rgb(path, width=800, height=416)
    before = path.read_bytes()

    result = preprocess_image_for_ocr(str(path), min_short_side=MIN_SHORT_SIDE)

    assert result is not None
    assert result.applied is True
    assert result.mime_type == "image/png"
    assert _png_size(result.image_data) == (1600, 832)
    assert path.read_bytes() == before
    assert result.image_data != before
    assert result.image_data.startswith(b"\x89PNG\r\n\x1a\n")


def test_min_short_side_zero_is_noop(tmp_path: Path) -> None:
    path = tmp_path / "tiny.png"
    _save_rgb(path, width=400, height=300)
    before = path.read_bytes()

    result = preprocess_image_for_ocr(str(path), min_short_side=0)

    assert result is not None
    assert result.applied is False
    assert result.image_data == before


def test_unreadable_file_returns_none(tmp_path: Path) -> None:
    path = tmp_path / "broken.png"
    path.write_bytes(b"\x89PNG\r\n\x1a\nfake")
    assert preprocess_image_for_ocr(str(path), min_short_side=MIN_SHORT_SIDE) is None


def test_pdf_like_file_returns_none(tmp_path: Path) -> None:
    path = tmp_path / "table.pdf"
    path.write_bytes(b"%PDF-1.4 not an image")
    assert preprocess_image_for_ocr(str(path), min_short_side=MIN_SHORT_SIDE) is None


def test_missing_file_returns_none(tmp_path: Path) -> None:
    assert preprocess_image_for_ocr(str(tmp_path / "nope.png"), min_short_side=MIN_SHORT_SIDE) is None


def test_rgba_small_image_converts_and_scales(tmp_path: Path) -> None:
    path = tmp_path / "rgba.png"
    Image.new("RGBA", (200, 100), color=(10, 20, 30, 255)).save(path, format="PNG")
    before = path.read_bytes()

    result = preprocess_image_for_ocr(str(path), min_short_side=MIN_SHORT_SIDE)

    assert result is not None
    assert result.applied is True
    width, height = _png_size(result.image_data)
    assert (width, height) == (400, 200)
    with Image.open(BytesIO(result.image_data)) as image:
        assert image.mode == "RGB"
    assert path.read_bytes() == before


def test_palette_small_image_converts_and_scales(tmp_path: Path) -> None:
    path = tmp_path / "palette.png"
    Image.new("RGB", (120, 80), color=(40, 80, 120)).convert("P").save(path, format="PNG")

    result = preprocess_image_for_ocr(str(path), min_short_side=MIN_SHORT_SIDE)

    assert result is not None
    assert result.applied is True
    assert _png_size(result.image_data) == (240, 160)


def test_oversized_png_encode_falls_back_to_original(tmp_path: Path) -> None:
    path = tmp_path / "small.png"
    _save_rgb(path, width=400, height=300)
    before = path.read_bytes()
    huge = b"\x89PNG\r\n\x1a\n" + (b"x" * (8 * 1024 * 1024 + 1))

    with patch("core.ocr.image_preprocess._encode_png", return_value=huge):
        result = preprocess_image_for_ocr(str(path), min_short_side=MIN_SHORT_SIDE)

    assert result is not None
    assert result.applied is False
    assert result.image_data == before
    assert path.read_bytes() == before
