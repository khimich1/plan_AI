#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тестовый скрипт для проверки оптимизации на конкретном заказе
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

# Твой заказ
orders_2d = [
    {'length': 3.31, 'width': 320, 'qty': 3},   # ПБ 33,1-3.2-8п
    {'length': 6.63, 'width': 860, 'qty': 4},   # ПБ 66,3-8,6-8п
    {'length': 5.6, 'width': 860, 'qty': 5},    # ПБ 56-8,6-8п
    {'length': 5.6, 'width': 320, 'qty': 11},   # ПБ 56-3,2-8п
]

print("="*70)
print("ТЕСТИРОВАНИЕ ОПТИМИЗАЦИИ")
print("="*70)
print("\nЗАКАЗ:")
total_requested = 0
for order in orders_2d:
    print(f"  • {order['qty']}× плита {order['length']}м × {order['width']}мм")
    total_requested += order['qty']
print(f"\nВСЕГО ЗАКАЗАНО: {total_requested} плит")

# Запускаем оптимизацию
result = optimize_with_cascading_longitudinal_cuts(orders_2d=orders_2d)

print("\n" + "="*70)
print("РЕЗУЛЬТАТ ОПТИМИЗАЦИИ:")
print("="*70)

if not result:
    print("❌ ОШИБКА: Оптимизация не вернула результат!")
else:
    print(f"\n📦 ВСЕГО ИСХОДНЫХ ПЛИТ 1200мм: {result.get('total_plates', 0)}")
    
    print(f"\n🔪 ПЕРВИЧНЫЕ РЕЗЫ ({len(result.get('primary_cuts', []))}):")
    for i, cut in enumerate(result.get('primary_cuts', []), 1):
        print(f"  {i}. Ширина {cut['width']}мм, остаток {cut['rest']}мм, кол-во {cut['qty']}")
        lengths = cut.get('lengths', [])
        if lengths:
            print(f"     Длины: {lengths[:5]}{'...' if len(lengths) > 5 else ''}")
        else:
            print(f"     Длины: НЕТ ДАННЫХ")
    
    print(f"\n🔄 ВТОРИЧНЫЕ РЕЗЫ ({len(result.get('secondary_cuts', []))}):")
    for i, sec in enumerate(result.get('secondary_cuts', []), 1):
        print(f"  {i}. Остаток {sec['source']}мм → {sec.get('pieces', 1)}× {sec['cuts'][0] if sec.get('cuts') else '?'}мм, кол-во {sec['qty']}")
        source_lengths = sec.get('source_lengths', [])
        if source_lengths:
            print(f"     Исходные длины остатков: {source_lengths[:5]}{'...' if len(source_lengths) > 5 else ''}")
        else:
            print(f"     Исходные длины: НЕТ ДАННЫХ")
        target_lengths = sec.get('lengths', [])
        if target_lengths:
            print(f"     Результирующие длины: {target_lengths[:5]}{'...' if len(target_lengths) > 5 else ''}")
        else:
            print(f"     Результирующие длины: НЕТ ДАННЫХ")
    
    print(f"\n📋 РАСПРЕДЕЛЕНИЕ ГОТОВЫХ ПЛИТ ({len(result.get('plate_assignments', []))}):")
    from collections import Counter
    assignments = result.get('plate_assignments', [])
    
    # Группируем по (length, width)
    plate_counts = Counter((p['length'], p['width']) for p in assignments)
    print("\nПолученные плиты:")
    total_produced = 0
    for (length, width), qty in sorted(plate_counts.items()):
        print(f"  • {qty}× плита {length}м × {width}мм")
        total_produced += qty
    
    print(f"\nВСЕГО ПРОИЗВЕДЕНО: {total_produced} плит")
    
    # Проверка соответствия заказу
    print("\n" + "="*70)
    print("ПРОВЕРКА СООТВЕТСТВИЯ ЗАКАЗУ:")
    print("="*70)
    
    for order in orders_2d:
        key = (order['length'], order['width'])
        produced = plate_counts.get(key, 0)
        status = "✅ OK" if produced >= order['qty'] else "❌ НЕДОСТАТОЧНО"
        print(f"  {order['length']}м × {order['width']}мм: заказано {order['qty']}, произведено {produced} {status}")
    
    if total_produced != total_requested:
        print(f"\n⚠️  ВНИМАНИЕ: Произведено {total_produced} плит, а заказано {total_requested}!")
        print(f"    Разница: {total_produced - total_requested:+d} плит")

print("\n" + "="*70)

