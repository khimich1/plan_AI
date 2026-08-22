#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from pathlib import Path

from PIL import Image

from core.ocr.image_meta import image_short_side_px


def _save_rgb(path: Path, width: int, height: int) -> None:
    Image.new("RGB", (width, height), color=(255, 255, 255)).save(path, format="PNG")


def test_image_short_side_px_returns_min_dimension(tmp_path: Path) -> None:
    path = tmp_path / "wide.png"
    _save_rgb(path, width=1200, height=800)
    assert image_short_side_px(str(path)) == 800


def test_image_short_side_px_portrait(tmp_path: Path) -> None:
    path = tmp_path / "tall.png"
    _save_rgb(path, width=400, height=900)
    assert image_short_side_px(str(path)) == 400


def test_image_short_side_px_unreadable_returns_none(tmp_path: Path) -> None:
    path = tmp_path / "broken.png"
    path.write_bytes(b"\x89PNG\r\n\x1a\nfake")
    assert image_short_side_px(str(path)) is None


def test_image_short_side_px_missing_file_returns_none(tmp_path: Path) -> None:
    assert image_short_side_px(str(tmp_path / "nope.png")) is None


def test_image_short_side_px_does_not_rewrite_file(tmp_path: Path) -> None:
    path = tmp_path / "keep.png"
    _save_rgb(path, width=320, height=240)
    before = path.read_bytes()
    assert image_short_side_px(str(path)) == 240
    assert path.read_bytes() == before
