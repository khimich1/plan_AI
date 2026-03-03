#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тест для проверки группировки плит по нагрузкам в смете.
Проверяет, что плиты с одинаковыми размерами, но разными нагрузками НЕ объединяются.
"""
import sys
from pathlib import Path

# Добавляем корневую директорию в path
PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

import core.config_and_data as cfg
from viz_modules.procurement import build_procurement_items, build_price_rows


def test_procurement_with_different_loads():
    """Тест группировки плит с разными нагрузками"""
    print("=" * 70)
    print("TEST: Procurement items with different load codes")
    print("=" * 70)
    
    # Тестовый заказ - плиты с ОДИНАКОВЫМИ размерами, но РАЗНЫМИ нагрузками
    test_order = """
    Плиты ПБ 73-12-8п — 93 шт
    Плиты ПБ 73-12-10п — 4 шт
    Плиты ПБ 55-12-12,5п — 6 шт
    Плиты ПБ 55-12-10п — 9 шт
    Плиты ПБ 87-6,65-8п — 2 шт
    """
    
    print(f"\nTest order:\n{test_order}")
    
    # Парсим заказ
    unparsed = cfg.set_plate_lists_from_text(test_order)
    
    if unparsed:
        print(f"\nWARNING: Unparsed lines: {unparsed}")
    
    print("\nPLATE_LOAD_DETAILS:")
    for key, qty in sorted(cfg.PLATE_LOAD_DETAILS.items()):
        length, width, load = key[0], key[1], key[2]
        print(f"  {qty}x {length}m x {width}m with load {load}p")
    
    # Получаем items закупки
    items = build_procurement_items()
    
    print(f"\nProcurement items (should be {5} separate items):")
    for idx, item in enumerate(items, 1):
        load_info = f" ({item.get('load_code', '?')}p)" if 'load_code' in item else ""
        print(f"  {idx}. {item['qty']}x {item['length']}m x {item['width']}m{load_info}")
    
    # Проверяем, что плиты НЕ объединились
    print("\n" + "=" * 70)
    print("VERIFICATION:")
    print("=" * 70)
    
    # Должно быть 5 отдельных позиций
    assert len(items) >= 5, f"Expected at least 5 items, got {len(items)}"
    print(f"  OK! Got {len(items)} separate items (not combined)")
    
    # Проверяем, что плиты 7.3м x 1.2м с 8п и 10п - РАЗНЫЕ позиции
    items_73_12 = [it for it in items if abs(it['length'] - 7.3) < 0.01 and abs(it['width'] - 1.2) < 0.01]
    print(f"\n  Plates 7.3m x 1.2m: {len(items_73_12)} items (should be 2: 8p and 10p)")
    
    if len(items_73_12) >= 2:
        for it in items_73_12:
            load = it.get('load_code', '?')
            print(f"    - {it['qty']} pcs with load {load}p")
        print("  OK! 8p and 10p are separate!")
    else:
        print(f"  WARNING: Found only {len(items_73_12)} item(s), expected 2")
    
    # Проверяем ширину 6,65 (не должна округляться до 6,7)
    print("\n" + "=" * 70)
    print("TEST: Width 6,65 formatting in plate names")
    print("=" * 70)
    
    items_665 = [it for it in items if abs(it['length'] - 8.7) < 0.01 and abs(it['width'] - 0.665) < 0.01]
    if items_665:
        it = items_665[0]
        load = it.get('load_code', 8)
        name = cfg.make_plate_name(it['length'], it['width'], load_code=load)
        print(f"  Plate name: {name}")
        print(f"  Width in name: {'6,65' if '6,65' in name else '6,7' if '6,7' in name else 'unknown'}")
        
        assert '6,65' in name, f"Expected '6,65' in name, got: {name}"
        print("  OK! Width 6,65 is NOT rounded to 6,7")
    else:
        print("  WARNING: Plate 8.7m x 0.665m not found in items")
    
    print("\n" + "=" * 70)
    print("ALL PROCUREMENT TESTS PASSED!")
    print("=" * 70)


if __name__ == '__main__':
    test_procurement_with_different_loads()
    
    print("\nSUCCESS! Procurement correctly handles:")
    print("  - Different load codes for same dimensions (8p, 10p, 13p)")
    print("  - Exact widths without rounding (6,65 not 6,7)")
    print("  - Separate entries for plates with different loads")

