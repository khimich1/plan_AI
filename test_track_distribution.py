#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тест: проверка распределения плит по дорожкам с учётом нагрузки
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core import config as cfg
from core.config import set_plate_lists_from_text
from core.optimization import optimize_with_cascading_longitudinal_cuts
from viz_modules.layout_sequence import build_layout_sequence
from collections import defaultdict

# Тестовый заказ
test_order = """
Плиты ПБ 73-12-8п — 93 шт
Плиты ПБ 73-12-10п — 4 шт
Плиты ПБ 55-12-12,5п — 6 шт
"""

print("=" * 100)
print("ТЕСТ РАСПРЕДЕЛЕНИЯ ПЛИТ ПО ДОРОЖКАМ С ГРУППИРОВКОЙ ПО НАГРУЗКАМ")
print("=" * 100)

# Парсим
unparsed_lines = set_plate_lists_from_text(test_order)
print(f"\n[1] Парсинг: найдено {len(cfg.PLATE_LOAD_DETAILS)} записей")

# Группируем по нагрузкам
orders_by_load = defaultdict(list)
for (length, width_m, load_code), qty in cfg.PLATE_LOAD_DETAILS.items():
    width_mm = int(round(width_m * 1000))
    orders_by_load[load_code].append({
        'length': length,
        'width': width_mm,
        'qty': qty,
        'load_code': load_code
    })

print(f"\n[2] Группировка: {len(orders_by_load)} групп")
for load_code in sorted(orders_by_load.keys()):
    orders = orders_by_load[load_code]
    total_qty = sum(o['qty'] for o in orders)
    print(f"  • {load_code}п: {total_qty} плит")

# Оптимизация
import core.optimization as optimization
optimization_results_by_load = {}

for load_code in sorted(orders_by_load.keys()):
    orders_2d = orders_by_load[load_code]
    optimization_result = optimize_with_cascading_longitudinal_cuts(orders_2d=orders_2d)
    if optimization_result and optimization_result.get('total_plates', 0) > 0:
        optimization_results_by_load[load_code] = optimization_result

optimization.OPT_CASCADING_PLAN_BY_LOAD = optimization_results_by_load

print(f"\n[3] Оптимизация: {len(optimization_results_by_load)} результатов")
for load_code in sorted(optimization_results_by_load.keys()):
    result = optimization_results_by_load[load_code]
    print(f"  • {load_code}п: {result['total_plates']} плит 1.2м")

# Построение последовательности
seq = build_layout_sequence()

print(f"\n[4] Построение последовательности:")
print(f"  Тип данных: {type(seq)}")
if isinstance(seq, list) and seq and isinstance(seq[0], dict):
    if 'load_code' in seq[0]:
        print(f"  ✅ НОВЫЙ ФОРМАТ: группировка по нагрузкам обнаружена!")
        print(f"  Групп: {len(seq)}")
        for group in seq:
            load_code = group['load_code']
            sequence = group['sequence']
            label = group.get('label', f'Нагрузка {load_code}п')
            print(f"\n  📦 {label}:")
            print(f"     Плит в группе: {len(sequence)}")
    else:
        print(f"  ⚠️ СТАРЫЙ ФОРМАТ: без группировки")
        print(f"     Плит: {len(seq)}")

# Разбиение на дорожки (симуляция логики из visualization.py)
MAX_TRACK_LENGTH = 101.0
tracks = []

if isinstance(seq, list) and seq and isinstance(seq[0], dict) and 'load_code' in seq[0]:
    print(f"\n[5] Разбиение на дорожки (НОВАЯ ЛОГИКА):")
    
    for group in seq:
        load_code = group['load_code']
        items = group['sequence']
        group_label = group.get('label', f'Нагрузка {load_code}п')
        
        print(f"\n  === {group_label} ===")
        
        current_track = []
        current_track_length = 0.0
        track_num = 1
        
        for item in items:
            item_length = item['length']
            
            if current_track_length + item_length > MAX_TRACK_LENGTH and current_track:
                tracks.append({
                    'items': current_track,
                    'length': current_track_length,
                    'load_code': load_code,
                    'label': group_label
                })
                print(f"    Дорожка {len(tracks)}: {len(current_track)} плит, {current_track_length:.2f}м, [{group_label}]")
                current_track = []
                current_track_length = 0.0
                track_num += 1
            
            current_track.append(item)
            current_track_length += item_length
        
        if current_track:
            tracks.append({
                'items': current_track,
                'length': current_track_length,
                'load_code': load_code,
                'label': group_label
            })
            print(f"    Дорожка {len(tracks)}: {len(current_track)} плит, {current_track_length:.2f}м, [{group_label}]")

print(f"\n{'='*100}")
print(f"📊 ИТОГОВОЕ РАСПРЕДЕЛЕНИЕ ПО ДОРОЖКАМ")
print(f"{'='*100}")
print(f"\nВсего дорожек: {len(tracks)}\n")

for i, track in enumerate(tracks, start=1):
    load_code = track.get('load_code', '?')
    label = track.get('label', 'N/A')
    length = track.get('length', 0)
    items_count = len(track.get('items', []))
    
    print(f"Дорожка {i:2d}: {label:20s} | {items_count:3d} плит | {length:6.2f}м")

# Проверка: каждая нагрузка должна быть в отдельных дорожках
print(f"\n{'='*100}")
print(f"✅ ПРОВЕРКА РАЗДЕЛЕНИЯ ПО НАГРУЗКАМ")
print(f"{'='*100}")

loads_in_tracks = defaultdict(list)
for i, track in enumerate(tracks, start=1):
    load_code = track.get('load_code')
    loads_in_tracks[load_code].append(i)

all_ok = True
for load_code, track_nums in sorted(loads_in_tracks.items()):
    print(f"\n{load_code}п: дорожки {track_nums}")
    
    # Проверяем, что дорожки одной нагрузки идут подряд (непрерывный блок)
    if track_nums == list(range(min(track_nums), max(track_nums) + 1)):
        print(f"  ✅ OK: дорожки идут подряд (непрерывный блок)")
    else:
        print(f"  ⚠️ ВНИМАНИЕ: дорожки НЕ идут подряд (плиты перемешаны)")
        all_ok = False

print(f"\n{'='*100}")
if all_ok:
    print("✅ ✅ ✅ ТЕСТ ПРОЙДЕН! Плиты правильно разделены по нагрузкам!")
else:
    print("❌ ТЕСТ НЕ ПРОЙДЕН! Плиты с разными нагрузками перемешаны!")
print(f"{'='*100}")

