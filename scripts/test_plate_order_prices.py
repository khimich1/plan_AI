#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Проверка цен по списку марок: парсинг, ключ length_dm, цена из pb.db / XLSX,
базовая цена с пересчётом на ширину (как в build_price_rows) и оценка продольного реза.

Запуск из корня проекта:
  python scripts/test_plate_order_prices.py
"""
from __future__ import annotations

import io
import sys
from pathlib import Path

# Windows-консоль часто cp1251; без этого падает print с «₽» и длинными тире
if hasattr(sys.stdout, "buffer"):
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    except Exception:
        pass

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import core.config_and_data as cfg
from core.plate_length_rules import resolve_length_dm
from core.price_db import get_price
from viz_modules.price_utils import find_price_for_plate, load_price_table_from_xlsx

# Заказ для проверки (марка, кол-во)
ORDER_LINES: list[tuple[str, int]] = [
    ("Плиты ПБ 38-2,6-8п", 1),
    ("Плиты ПБ 41-2,6-8п", 1),
    ("Плиты ПБ 37-3,2-8п", 1),
    ("Плиты ПБ 47-3,2-8п", 1),
    ("Плиты ПБ 57-3,2-8п", 1),
    ("Плиты ПБ 29-4,6-8п", 1),
    ("Плиты ПБ 23-12-8п", 3),
    ("Плиты ПБ 29-12-8п", 3),
    ("Плиты ПБ 37-12-8п", 2),
    ("Плиты ПБ 41-12-8п", 3),
    ("Плиты ПБ 47-12-8п", 5),
    ("Плиты ПБ 57-12-8п", 2),
]


def _long_cuts_for_width(width_m: float) -> int:
    if abs(width_m - 1.2) < 0.01:
        return 0
    return 1 if width_m < 1.15 else 0


def main() -> None:
    price_table = load_price_table_from_xlsx(str(cfg.PRICE_XLSX_PATH))
    db_path = str(cfg.PRICE_DB_PATH)

    print("Прайс XLSX:", "OK" if price_table else "пусто / не найден")
    print("pb.db:", db_path, "(exists)" if Path(db_path).exists() else "(нет файла)")
    print()

    header = (
        f"{'Марка':<32} {'Lм':>6} {'Wм':>6} {'ldm':>4} {'lc':>3} "
        f"{'rub_1.2m':>12} {'base_1pc':>12} {'cut_1pc':>10} {'unit':>12} {'qty':>4} {'line_sum':>14}"
    )
    print(header)
    print("-" * len(header))

    total_sum = 0.0
    for name, qty in ORDER_LINES:
        L, W = cfg.parse_name_to_sizes(name)
        if L is None or W is None:
            print(f"{name[:32]:<32} ОШИБКА парсинга")
            continue

        load_code = cfg.parse_load_code_from_name(name, default=8)
        ldm = resolve_length_dm(L)
        db_price = get_price(L, load_code, db_path)
        use_fallback = db_price is None or (isinstance(db_price, (int, float)) and db_price <= 0)
        xlsx_price = find_price_for_plate(price_table, L, load_code) if use_fallback else None
        base_1_2 = (db_price if (db_price is not None and float(db_price) > 0) else None) or xlsx_price or 0.0

        if base_1_2 > 0:
            base_unit = base_1_2 * (W / 1.2)
        else:
            base_unit = 0.0

        lcuts = _long_cuts_for_width(W)
        if lcuts and qty > 0:
            long_cut_per_plate = lcuts * cfg.LONG_CUT_PRICE_PER_M * L
        else:
            long_cut_per_plate = 0.0

        unit_total = base_unit + long_cut_per_plate
        line_sum = unit_total * qty
        total_sum += line_sum

        short_name = name.replace("Плиты ", "")[:32]
        print(
            f"{short_name:<32} {L:>6.3f} {W:>6.3f} {ldm:>4} {load_code:>3} "
            f"{base_1_2:>12,.2f} {base_unit:>12,.2f} {long_cut_per_plate:>10,.2f} "
            f"{unit_total:>12,.2f} {qty:>4} {line_sum:>14,.2f}"
        )

    print("-" * len(header))
    print(f"{'ИТОГО (оценка без поперечных резов и плана обрезков)':>80} {total_sum:>14,.2f}")


if __name__ == "__main__":
    main()
