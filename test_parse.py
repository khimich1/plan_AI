#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тестируем парсинг заказа
"""

import sys
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except:
        import codecs
        sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')

import config_and_data as cfg
from config_and_data import set_plate_lists_from_text

# Твой текст заказа (ТОЧНО как ты отправляешь боту)
order_text = """
Плиты ПБ 33,1-3.2-8п 3
Плиты ПБ 66,3-8,6-8п 4
Плиты ПБ 56-8,6-8п 5
Плиты ПБ 56-3,2-8п 11
"""

print("="*70)
print("ТЕСТ ПАРСИНГА ЗАКАЗА")
print("="*70)
print("\nТЕКСТ ЗАКАЗА:")
print(order_text)

# Парсим
set_plate_lists_from_text(order_text)

print("\n" + "="*70)
print("РЕЗУЛЬТАТ ПАРСИНГА:")
print("="*70)

print(f"\nПЛАТЫ 320мм: {cfg.PLATES_0_32}")
print(f"ПЛИТЫ 860мм: {cfg.PLATES_0_86}")
print(f"ПЛИТЫ 1200мм: {cfg.PLATES_1_2}")

# Подсчитываем
from collections import Counter
counter_320 = Counter(cfg.PLATES_0_32)
counter_860 = Counter(cfg.PLATES_0_86)

print("\n" + "="*70)
print("АНАЛИЗ:")
print("="*70)

print("\nПлиты 320мм по длинам:")
for length, qty in sorted(counter_320.items()):
    print(f"  • {length}м: {qty} шт")
print(f"ВСЕГО 320мм: {len(cfg.PLATES_0_32)} шт")

print("\nПлиты 860мм по длинам:")
for length, qty in sorted(counter_860.items()):
    print(f"  • {length}м: {qty} шт")
print(f"ВСЕГО 860мм: {len(cfg.PLATES_0_86)} шт")

print(f"\nПлиты 1200мм: {len(cfg.PLATES_1_2)} шт")

print("\n" + "="*70)
print("ОЖИДАЛОСЬ:")
print("="*70)
print("  • 320мм: 3× 3.31м + 11× 5.6м = 14 шт")
print("  • 860мм: 4× 6.63м + 5× 5.6м = 9 шт")
print("  • 1200мм: 0 шт")

total_expected = 14 + 9
total_got = len(cfg.PLATES_0_32) + len(cfg.PLATES_0_86) + len(cfg.PLATES_1_2)

if total_got == total_expected:
    print(f"\n✅ ПАРСИНГ ПРАВИЛЬНЫЙ! Получено {total_got} плит")
else:
    print(f"\n❌ ОШИБКА ПАРСИНГА! Ожидалось {total_expected} плит, получено {total_got}")

print("\n" + "="*70)

