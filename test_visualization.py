#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тестовый скрипт для проверки визуализации
"""

import sys

# Настройка кодировки для Windows консоли
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except:
        import codecs
        sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')

from optimization import optimize_with_cascading_longitudinal_cuts
import optimization

# Твой заказ
orders_2d = [
    {'length': 3.31, 'width': 320, 'qty': 3},   # ПБ 33,1-3.2-8п
    {'length': 6.63, 'width': 860, 'qty': 4},   # ПБ 66,3-8,6-8п
    {'length': 5.6, 'width': 860, 'qty': 5},    # ПБ 56-8,6-8п
    {'length': 5.6, 'width': 320, 'qty': 11},   # ПБ 56-3,2-8п
]

print("="*70)
print("ТЕСТ ВИЗУАЛИЗАЦИИ")
print("="*70)

# Запускаем оптимизацию
result = optimize_with_cascading_longitudinal_cuts(orders_2d=orders_2d)

# Сохраняем в глобальную переменную
optimization.OPT_CASCADING_PLAN = result

print("\nОптимизация завершена. Теперь проверим визуализацию...")

# Импортируем визуализацию
from visualization import build_layout_sequence

print("\n" + "="*70)
print("РЕЗУЛЬТАТ build_layout_sequence():")
print("="*70)

sequence = build_layout_sequence()

print(f"\nВсего сегментов в последовательности: {len(sequence)}")
print(f"\nДетализация сегментов:\n")

for i, seg in enumerate(sequence, 1):
    mode = seg.get('mode', 'unknown')
    length = seg.get('length', 0)
    
    if mode == 'solid':
        label = seg.get('label', '?')
        print(f"{i}. SOLID: {length}м — {label}")
    
    elif mode == 'transverse':
        target = seg.get('target_length', 0)
        remainder = seg.get('remainder', 0)
        label = seg.get('label_target', '?')
        print(f"{i}. TRANSVERSE: {length}м → {target}м (остаток {remainder:.2f}м) — {label}")
    
    elif mode == 'split':
        main_w = seg.get('main_w', 0)
        rest_w = seg.get('rest_w', 0)
        label_main = seg.get('label_main', '?')
        label_rest = seg.get('label_rest', '?')
        secondary = seg.get('secondary_cuts', [])
        
        print(f"{i}. SPLIT: {length}м, основная {main_w*1000:.0f}мм, остаток {rest_w*1000:.0f}мм")
        print(f"   Основная: {label_main}")
        
        if secondary:
            print(f"   Вторичные резы ({len(secondary)}):")
            for j, sec in enumerate(secondary, 1):
                sec_label = sec.get('label', '?')
                sec_width = sec.get('width', 0)
                has_transverse = sec.get('has_transverse', False)
                target_length = sec.get('target_length')
                print(f"     {j}) Ширина {sec_width*1000:.0f}мм: {sec_label}", end='')
                if has_transverse:
                    print(f" [ПОПЕРЕЧНЫЙ РЕЗ → {target_length}м]")
                else:
                    print()
        elif label_rest and label_rest != '?':
            print(f"   Остаток: {label_rest}")

print("\n" + "="*70)
print("АНАЛИЗ:")
print("="*70)

# Подсчитываем плиты "О ПБ 56-3,2-8п"
count_o_56_32 = 0
for seg in sequence:
    if seg.get('mode') == 'split':
        secondary = seg.get('secondary_cuts')
        if secondary:  # Проверяем, что не None
            for sec in secondary:
                label = sec.get('label', '')
                if 'О ПБ 56-3,2-8п' in label or 'О ПБ 56-3.2-8п' in label:
                    count_o_56_32 += 1

print(f"\nНайдено сегментов 'О ПБ 56-3,2-8п': {count_o_56_32}")
print(f"\nПо плану должно быть:")
print(f"  • Из остатка 880мм (3.31м): 2 куска × 1 остаток = 2")
print(f"  • Из остатка 340мм (6.63м): 1 кусок × 4 остатка = 4")
print(f"  • Из остатка 340мм (5.6м): 1 кусок × 4 остатка = 4")
print(f"  • Из остатка 880мм (5.6м): 2 куска × 1 остаток = 2")
print(f"  ИТОГО: 12 сегментов с 'О ПБ 56-3,2-8п'")

if count_o_56_32 > 12:
    print(f"\n⚠️  ПРОБЛЕМА: Показано {count_o_56_32} сегментов, а должно быть 12!")
    print(f"    Визуализация дублирует вторичные резы!")
elif count_o_56_32 == 12:
    print(f"\n✅ ВСЁ ПРАВИЛЬНО: Количество сегментов соответствует ожиданию")
else:
    print(f"\n❓ Показано меньше сегментов: {count_o_56_32} вместо 12")

print("\n" + "="*70)

