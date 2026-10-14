"""Classify a factory price workbook by filename anchors (never by row count)."""

from __future__ import annotations

import os
import re
from typing import Literal

ProductKind = Literal[
    "composite_pile",
    "plates",
    "fbs",
    "step",
    "bridge_pile",
    "pile",
    "march",
]

MSG_UNKNOWN_KIND = "не понял группу по имени файла"

_FILENAME_DATE_RE = re.compile(r"(\d{2})[.\-](\d{2})[.\-](\d{2,4})")


class ClassifyError(ValueError):
    """Filename has no product-group anchor."""


def _normalize_filename(filename: str) -> str:
    base = os.path.basename(str(filename or "")).strip().lower().replace("ё", "е")
    return base


def classify_price_filename(filename: str) -> ProductKind:
    """Return product kind. First matching anchor wins. Raises ClassifyError if none."""
    name = _normalize_filename(filename)
    if "составн" in name:
        return "composite_pile"
    if "расчет новых цен на пб" in name or ("расчет" in name and "пб" in name):
        return "plates"
    if "фбс" in name:
        return "fbs"
    if "ступен" in name:
        return "step"
    if "мостов" in name:
        return "bridge_pile"
    if "цельн" in name:
        return "pile"
    if "лм" in name:
        return "march"
    raise ClassifyError(MSG_UNKNOWN_KIND)


def parse_price_list_date_from_filename(filename: str) -> str | None:
    """Extract YYYY-MM-DD from a ``dd.mm.yyyy`` (or yy) stamp in the basename."""
    match = _FILENAME_DATE_RE.search(os.path.basename(str(filename or "")))
    if not match:
        return None
    day, month, year = match.groups()
    if len(year) == 2:
        year = f"20{year}"
    return f"{year}-{month}-{day}"
