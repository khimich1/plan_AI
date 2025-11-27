#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тест для проверки работы с нагрузками (8п, 10п, 12,5п).
Проверяет, что:
1. Нагрузки правильно парсятся и сохраняются
2. Ширины 6,65 дм не округляются до 6,7
3. Плиты с одинаковыми размерами, но разными нагрузками различаются
"""
import sys
from pathlib import Path

# Добавляем корневую директорию в path
PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

import core.config_and_data as cfg


def test_load_codes_and_widths():
    """Тест парсинга нагрузок и точных ширин"""
    print("=" * 70)
    print("TEST: Load codes (8п/10п/12,5п) and exact widths (6,65)")
    print("=" * 70)
    
    # Тестовый заказ - плиты с РАЗНЫМИ нагрузками
    test_order = """
    Плиты ПБ 73-12-8п — 93 шт
    Плиты ПБ 73-12-10п — 4 шт
    Плиты ПБ 87-6,65-8п — 2 шт
    Плиты ПБ 55-12-12,5п — 6 шт
    """
    
    print(f"\nTest order:\n{test_order}")
    
    # Парсим заказ
    unparsed = cfg.set_plate_lists_from_text(test_order)
    
    if unparsed:
        print(f"\nWARNING: Unparsed lines: {unparsed}")
    else:
        print("\nOK: All lines parsed successfully!")
    
    # Проверяем PLATE_LOAD_DETAILS
    print("\nPLATE_LOAD_DETAILS (new format with load codes):")
    for key, qty in sorted(cfg.PLATE_LOAD_DETAILS.items()):
        length, width, load = key
        print(f"  {qty}x plate {length}m x {width}m with load {load}п ({int(width*1000)}mm)")
    
    # Тест 1: Разные нагрузки для одинаковых размеров
    print("\n" + "=" * 70)
    print("TEST 1: Different load codes for same dimensions")
    print("=" * 70)
    
    # Плиты 7.3м × 1.2м: 93 с нагрузкой 8п + 4 с нагрузкой 10п
    qty_8p = cfg.PLATE_LOAD_DETAILS.get((7.3, 1.2, 8), 0)
    qty_10p = cfg.PLATE_LOAD_DETAILS.get((7.3, 1.2, 10), 0)
    
    print(f"  Plates 7.3m x 1.2m with 8п: {qty_8p} pcs (expected: 93)")
    print(f"  Plates 7.3m x 1.2m with 10п: {qty_10p} pcs (expected: 4)")
    
    assert qty_8p == 93, f"Expected 93 plates with 8п, got {qty_8p}"
    assert qty_10p == 4, f"Expected 4 plates with 10п, got {qty_10p}"
    print("  OK! Different load codes are stored separately")
    
    # Тест 2: Нагрузка 12,5п
    print("\n" + "=" * 70)
    print("TEST 2: Load code 12,5п (should be rounded to 13)")
    print("=" * 70)
    
    qty_12_5p = cfg.PLATE_LOAD_DETAILS.get((5.5, 1.2, 13), 0)  # 12.5 округляется до 13
    print(f"  Plates 5.5m x 1.2m with 13п (12,5п): {qty_12_5p} pcs (expected: 6)")
    
    assert qty_12_5p == 6, f"Expected 6 plates with 13п, got {qty_12_5p}"
    print("  OK! Load code 12,5п correctly parsed and rounded to 13")
    
    # Тест 3: Точная ширина 6,65 (не должна округляться до 6,7)
    print("\n" + "=" * 70)
    print("TEST 3: Exact width 6,65 dm (should NOT be rounded to 6,7)")
    print("=" * 70)
    
    exact_width = cfg.get_exact_width(8.7, 'PLATES_0_70', 0.70)
    print(f"  Plate 8.7m exact width: {exact_width}m ({int(exact_width*1000)}mm)")
    print(f"  Expected: 0.665m (665mm), got: {exact_width}m")
    
    assert abs(exact_width - 0.665) < 0.001, f"Expected 0.665m, got {exact_width}m"
    print("  OK! Width 6,65 is NOT rounded (665mm, not 670mm)")
    
    # Тест 4: Получение нагрузки через get_load_code_for_plate
    print("\n" + "=" * 70)
    print("TEST 4: Getting correct load code via get_load_code_for_plate()")
    print("=" * 70)
    
    # Для плит 7.3м × 1.2м должна вернуться нагрузка 8п (их больше: 93 vs 4)
    load_73_12 = cfg.get_load_code_for_plate(7.3, 1.2, default=8)
    print(f"  Load for 7.3m x 1.2m: {load_73_12}п (expected: 8п - most common)")
    assert load_73_12 == 8, f"Expected 8п (most common), got {load_73_12}п"
    print("  OK! Returns most common load code (8п) for 7.3m x 1.2m")
    
    # Для плит 5.5м × 1.2м должна вернуться нагрузка 13п
    load_55_12 = cfg.get_load_code_for_plate(5.5, 1.2, default=8)
    print(f"  Load for 5.5m x 1.2m: {load_55_12}п (expected: 13п)")
    assert load_55_12 == 13, f"Expected 13п, got {load_55_12}п"
    print("  OK! Returns correct load code (13п) for 5.5m x 1.2m")
    
    print("\n" + "=" * 70)
    print("ALL TESTS PASSED!")
    print("=" * 70)


def test_width_formatting():
    """Тест форматирования ширин (6.65 → "6,65", 5.3 → "5,3", 12.0 → "12")"""
    print("\n" + "=" * 70)
    print("TEST: Width formatting (умное форматирование)")
    print("=" * 70)
    
    test_cases = [
        (6.65, "6,65"),   # Два знака
        (5.3, "5,3"),     # Один знак
        (12.0, "12"),     # Целое число
        (7.2, "7,2"),     # Один знак
        (10.2, "10,2"),   # Один знак
    ]
    
    for width_dm, expected in test_cases:
        # Применяем умное форматирование
        result = f"{width_dm:.2f}".rstrip('0').rstrip('.').replace('.', ',')
        print(f"  {width_dm} dm -> \"{result}\" (expected: \"{expected}\")")
        assert result == expected, f"Expected {expected}, got {result}"
    
    print("\n  OK! All width formatting tests passed")


if __name__ == '__main__':
    test_load_codes_and_widths()
    test_width_formatting()
    
    print("\nSUCCESS! All load code and width formatting tests passed!")
    print("The bot now correctly handles:")
    print("  - Different load codes (8п, 10п, 12,5п)")
    print("  - Exact widths without rounding (6,65 not 6,7)")
    print("  - Multiple plates with same dimensions but different loads")

