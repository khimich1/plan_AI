#!/usr/bin/env python3
"""Remove per-function ensure_schema / _init_schema (Stage 5d A4 cleanup)."""

from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

FOR_OFFERS_MANAGERS = [
    ROOT / "core" / "kp_db_offers.py",
    ROOT / "core" / "kp_db_managers.py",
]

FOR_PLATES_RESTS = [
    ROOT / "core" / "kp_db_plates.py",
    ROOT / "core" / "kp_db_rests.py",
]


def strip_ensure_schema(path: Path) -> None:
    text = path.read_text(encoding="utf-8")
    text = re.sub(r"from core\.kp_db_schema import ensure_schema\n\n?", "", text)
    text = re.sub(r"    ensure_schema\(db_path\)\n\n?", "", text)
    text = re.sub(r"    ensure_schema\(db_path\)\n", "", text)
    path.write_text(text, encoding="utf-8")
    print("cleaned ensure_schema in", path.name)


def strip_init_schema(path: Path) -> None:
    text = path.read_text(encoding="utf-8")
    text = re.sub(
        r"def _init_schema\(db_path: str\) -> None:\n"
        r"    from core\.kp_db import ensure_schema\n\n?"
        r"    ensure_schema\(db_path\)\n\n?",
        "",
        text,
    )
    text = re.sub(r"        _init_schema\(db_path\)\n\n?", "", text)
    text = re.sub(r"    _init_schema\(db_path\)\n\n?", "", text)
    text = re.sub(r"        _init_schema\(db_path\)\n", "", text)
    text = re.sub(r"    _init_schema\(db_path\)\n", "", text)
    path.write_text(text, encoding="utf-8")
    print("cleaned _init_schema in", path.name)


def main() -> None:
    for p in FOR_OFFERS_MANAGERS:
        strip_ensure_schema(p)
    for p in FOR_PLATES_RESTS:
        strip_init_schema(p)


if __name__ == "__main__":
    main()
