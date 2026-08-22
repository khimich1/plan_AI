#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Optional 2× Lanczos + autocontrast for small OCR frames. Does not rewrite the file."""

from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path

from PIL import Image, ImageOps, UnidentifiedImageError

from core.ocr.providers.openai import image_mime_type

_MAX_PROCESSED_PNG_BYTES = 8 * 1024 * 1024


@dataclass(frozen=True)
class OcrPreprocessResult:
    image_data: bytes
    mime_type: str
    applied: bool


def _encode_png(image: Image.Image) -> bytes:
    buffer = BytesIO()
    image.save(buffer, format="PNG")
    return buffer.getvalue()


def preprocess_image_for_ocr(
    image_path: str,
    *,
    min_short_side: int,
) -> OcrPreprocessResult | None:
    """Return payload bytes for the vision API, or None if the frame cannot be read.

    applied=False keeps the original bytes (threshold not taken, min_short_side==0,
    or encoded PNG larger than 8 MiB).
    """
    try:
        original = Path(image_path).read_bytes()
    except OSError:
        return None
    if not original:
        return None

    try:
        with Image.open(image_path) as image:
            width, height = image.size
            if width <= 0 or height <= 0:
                return None
            working = image.convert("RGB")
    except (OSError, UnidentifiedImageError, ValueError):
        return None

    original_mime = image_mime_type(image_path, original)
    short_side = min(width, height)
    if min_short_side <= 0 or short_side >= min_short_side:
        return OcrPreprocessResult(original, original_mime, False)

    scaled = working.resize((width * 2, height * 2), resample=Image.Resampling.LANCZOS)
    contrasted = ImageOps.autocontrast(scaled, cutoff=1)
    encoded = _encode_png(contrasted)
    if len(encoded) > _MAX_PROCESSED_PNG_BYTES:
        return OcrPreprocessResult(original, original_mime, False)
    return OcrPreprocessResult(encoded, "image/png", True)
