#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Быстрая smoke-проверка проекта.

Запускать из корня проекта:
    python scripts/smoke_check.py

Проверяет:
- что базы данных существуют (если нет — предупреждает)
- что парсер плит работает на паре примеров

Telegram-бот soft-decommissioned (P5 WP1): см. bot/README.md.
"""

from __future__ import annotations

import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

if sys.platform == "win32":
    # На Windows консоль часто cp1251, она не умеет часть символов Unicode.
    # Включаем UTF-8, чтобы скрипт не падал на выводе.
    import io

    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")


def main() -> int:
    print("=" * 70)
    print("SMOKE CHECK")
    print("=" * 70)
    print(f"Project root: {PROJECT_ROOT}")
    print("[INFO] Telegram bot deprecated — see bot/README.md")

    from core.project_paths import PRICE_DB_PATH

    if PRICE_DB_PATH.exists():
        print(f"[OK] Найдена база цен pb.db: {PRICE_DB_PATH}")
    else:
        print(f"[WARN] Не найдена база цен pb.db: {PRICE_DB_PATH}")
        print("   - Это не всегда критично: прайс может подхватиться из Excel")

    plita_db = PROJECT_ROOT / "plita.db"
    if plita_db.exists():
        print(f"[OK] Найдена база КП plita.db: {plita_db}")
    else:
        print(f"[WARN] Не найдена база КП plita.db: {plita_db}")
        print("   - Она создастся при первом сохранении КП")

    from core.config_and_data import set_plate_lists_from_text

    examples = [
        "ПБ 78-12-8п 5 шт\nПБ 66,2-12-8п 6",
        "1.2×3.39 — 2 шт\n0,32x6,63 - 4",
    ]

    for i, text in enumerate(examples, 1):
        try:
            unparsed, _contributions, _line_loads = set_plate_lists_from_text(text)
            print(f"[OK] Парсер: пример #{i} — ok, нераспознано строк: {len(unparsed)}")
            if unparsed:
                for line in unparsed[:5]:
                    print(f"   - {line}")
        except Exception as e:
            print(f"[FAIL] Парсер: пример #{i} — ошибка: {e}")

    print("=" * 70)
    print("ГОТОВО")
    print("=" * 70)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
