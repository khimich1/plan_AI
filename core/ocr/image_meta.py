#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Read image pixel size without mutating the file."""

from __future__ import annotations

from typing import Optional

from PIL import Image, UnidentifiedImageError


def image_short_side_px(image_path: str) -> Optional[int]:
    """Return min(width, height) in pixels, or None if the frame cannot be read."""
    try:
        with Image.open(image_path) as image:
            width, height = image.size
    except (OSError, UnidentifiedImageError, ValueError):
        return None
    if width <= 0 or height <= 0:
        return None
    return min(width, height)
