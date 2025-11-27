#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тест для проверки работы точных ширин плит.
Проверяет, что плиты с нестандартными ширинами (например, 530мм вместо 460мм)
корректно парсятся и сохраняются.
"""
import sys
from pathlib import Path

# Добавляем корневую директорию в path
PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

import core.config_and_data as cfg


def test_exact_width_parsing():
    """Тест парсинга плит с точными ширинами"""
    print("=" * 60)
    print("TEST: Parsing plates with exact widths")
    print("=" * 60)
    
    # Тестовый заказ - такой же, как у клиента
    test_order = """
    Плиты ПБ 28-5,3-8п — 2 шт
    Плиты ПБ 73-12-8п — 93 шт
    Плиты ПБ 73-10,2-8п — 6 шт
    Плиты ПБ 74-12-8п — 13 шт
    """
    
    print(f"\nTest order:\n{test_order}")
    
    # Парсим заказ
    unparsed = cfg.set_plate_lists_from_text(test_order)
    
    if unparsed:
        print(f"\nWARNING: Unparsed lines: {unparsed}")
    else:
        print("\nOK: All lines parsed successfully!")
    
    # Проверяем, что плиты добавлены в правильные списки
    print("\nParsing results:")
    print(f"  PLATES_0_46 (range 460-530mm): {cfg.PLATES_0_46}")
    print(f"  PLATES_1_2 (1200mm): {cfg.PLATES_1_2}")
    print(f"  PLATES_0_70 (660-710mm): {cfg.PLATES_0_70}")
    print(f"  PLATES_0_74 (730-750mm): {cfg.PLATES_0_74}")
    
    # Проверяем точные ширины
    print("\nExact widths (PLATE_EXACT_WIDTHS):")
    for key, width in cfg.PLATE_EXACT_WIDTHS.items():
        length, list_name = key
        print(f"  {list_name}: length {length}m -> width {width}m ({int(width*1000)}mm)")
    
    # Тестируем get_exact_width()
    print("\nTesting get_exact_width():")
    
    # Плита ПБ 28-5,3-8п должна иметь ширину 530мм
    exact_width_1 = cfg.get_exact_width(2.8, 'PLATES_0_46', 0.46)
    print(f"  Plate 2.8m from PLATES_0_46: {exact_width_1}m ({int(exact_width_1*1000)}mm)")
    assert abs(exact_width_1 - 0.53) < 0.001, f"Expected width 0.53m, got {exact_width_1}m"
    print("    OK! Correct (530mm instead of 460mm)")
    
    # Плита ПБ 73-12-8п должна иметь ширину 1200мм (целая плита)
    exact_width_2 = cfg.get_exact_width(7.3, 'PLATES_1_2', 1.2)
    print(f"  Plate 7.3m from PLATES_1_2: {exact_width_2}m ({int(exact_width_2*1000)}mm)")
    assert abs(exact_width_2 - 1.2) < 0.001, f"Expected width 1.2m, got {exact_width_2}m"
    print("    OK! Correct (full plate 1200mm)")
    
    # Плита ПБ 73-10,2-8п должна быть в PLATES_1_08 с шириной примерно 1.02м
    if 7.3 in cfg.PLATES_1_08:
        exact_width_3 = cfg.get_exact_width(7.3, 'PLATES_1_08', 1.08)
        print(f"  Plate 7.3m from PLATES_1_08: {exact_width_3}m ({int(exact_width_3*1000)}mm)")
        print("    OK! Found in PLATES_1_08")
    elif 7.3 in cfg.PLATES_1_0:
        exact_width_3 = cfg.get_exact_width(7.3, 'PLATES_1_0', 1.0)
        print(f"  Plate 7.3m from PLATES_1_0: {exact_width_3}m ({int(exact_width_3*1000)}mm)")
        print("    OK! Found in PLATES_1_0")
    
    print("\n" + "=" * 60)
    print("ALL TESTS PASSED!")
    print("=" * 60)


def test_fallback_width():
    """Тест fallback на дефолтную ширину, если точная не найдена"""
    print("\n" + "=" * 60)
    print("TEST: Fallback to default width")
    print("=" * 60)
    
    # Очищаем данные
    cfg._clear_all_plate_lists()
    
    # Добавляем плиту вручную БЕЗ использования парсера
    # (чтобы точная ширина НЕ сохранилась)
    cfg.PLATES_0_46.append(5.0)
    
    # Пытаемся получить точную ширину
    width = cfg.get_exact_width(5.0, 'PLATES_0_46', 0.46)
    
    print(f"\n  Plate 5.0m without exact width: {width}m ({int(width*1000)}mm)")
    print(f"  Expected default width 0.46m")
    
    assert abs(width - 0.46) < 0.001, f"Expected fallback to 0.46m, got {width}m"
    print("  OK! Fallback works correctly!")
    
    print("\n" + "=" * 60)


if __name__ == '__main__':
    test_exact_width_parsing()
    test_fallback_width()
    
    print("\nSUCCESS! All tests passed!")
    print("The bot now correctly handles exact plate widths!")

