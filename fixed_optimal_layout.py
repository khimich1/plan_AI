#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Исправленная оптимальная раскладка плит
"""

import sqlite3
import pandas as pd
from collections import defaultdict
import math
import re

def load_and_parse_plates():
    """Загружает и парсит плиты из базы данных"""
    conn = sqlite3.connect('pb.db')
    df = pd.read_sql_query("SELECT * FROM plity_ex", conn)
    conn.close()
    
    plates = []
    
    for idx, row in df.iterrows():
        # Получаем маркировку
        marking = row['Маркировка плиты']
        
        if pd.notna(marking) and marking:
            # Парсим маркировку: "Плита ПБ 28-9,2-10п"
            pattern = r'(\d+[,.]?\d*)-(\d+[,.]?\d*)-(\d+)'
            match = re.search(pattern, marking)
            
            if match:
                length_dm = float(match.group(1).replace(',', '.'))
                width_dm = float(match.group(2).replace(',', '.'))
                load_code = int(match.group(3))
                
                length_m = length_dm / 10
                width_m = width_dm / 10
                load_kg = load_code * 100
                
                # Получаем неделю армирования (срочность)
                week_arm = row['Столбец армирования']
                
                # Определяем приоритет срочности
                if pd.isna(week_arm) or week_arm == '':
                    priority = 999  # самый низкий приоритет
                    urgency = "Нет недели армирования"
                else:
                    try:
                        priority = float(week_arm)
                        urgency = f"Неделя {priority}"
                    except:
                        priority = 999
                        urgency = "Неизвестная неделя"
                
                plate = {
                    'id': idx,
                    'marking': marking,
                    'length': length_m,
                    'width': width_m,
                    'load_capacity': load_kg,
                    'week_arm': week_arm,
                    'priority': priority,
                    'urgency': urgency,
                    'reinforcement': week_arm
                }
                plates.append(plate)
    
    return plates

def find_optimal_solution(plates, road_width=1.2, road_length=101):
    """Находит оптимальное решение"""
    
    # Сортируем плиты по приоритету (срочности)
    sorted_plates = sorted(plates, key=lambda x: x['priority'])
    
    print(f"\nПлиты отсортированы по срочности (первые 10):")
    for plate in sorted_plates[:10]:
        print(f"  Приоритет {plate['priority']}: {plate['marking']} ({plate['width']}x{plate['length']}м)")
    
    # Группируем по армированию
    plates_by_reinforcement = defaultdict(list)
    for plate in sorted_plates:
        plates_by_reinforcement[plate['reinforcement']].append(plate)
    
    best_solution = None
    min_waste = float('inf')
    
    print(f"\nПроверяем комбинации для дорожки {road_width}м x {road_length}м:")
    
    # Проверяем каждую группу армирования
    for reinforcement, group_plates in plates_by_reinforcement.items():
        print(f"\nАрмирование '{reinforcement}': {len(group_plates)} плит")
        
        # Сортируем внутри группы по приоритету
        group_plates.sort(key=lambda x: x['priority'])
        
        # Ищем плиты, подходящие для дорожки
        suitable_plates = []
        
        for plate in group_plates:
            # Плита точно подходит по ширине
            if abs(plate['width'] - road_width) < 0.01:
                suitable_plates.append(('exact_width', plate))
            # Плита меньше ширины дорожки
            elif plate['width'] < road_width:
                suitable_plates.append(('smaller', plate))
            # Плита больше ширины дорожки (нужен разрез)
            elif plate['width'] > road_width:
                suitable_plates.append(('larger', plate))
        
        if suitable_plates:
            # Берем самую срочную подходящую плиту
            plate_type, plate = suitable_plates[0]
            
            if plate_type == 'exact_width':
                # Плита точно подходит
                plates_needed = math.ceil(road_length / plate['length'])
                total_length = plates_needed * plate['length']
                
                used_area = plates_needed * plate['width'] * plate['length']
                road_area = road_width * road_length
                
                waste_percent = ((used_area - road_area) / road_area * 100) if road_area > 0 else 0
                
                solution = {
                    'reinforcement': reinforcement,
                    'priority': plate['priority'],
                    'plate_type': 'exact_width',
                    'plate': plate,
                    'num_plates': plates_needed,
                    'num_cuts': 0,
                    'waste_percent': waste_percent,
                    'total_length': total_length
                }
                
            elif plate_type == 'smaller':
                # Несколько плит поперек
                plates_across = math.ceil(road_width / plate['width'])
                plates_length = math.ceil(road_length / plate['length'])
                total_plates = plates_across * plates_length
                
                used_area = total_plates * plate['width'] * plate['length']
                road_area = road_width * road_length
                
                waste_percent = ((used_area - road_area) / road_area * 100) if road_area > 0 else 0
                
                solution = {
                    'reinforcement': reinforcement,
                    'priority': plate['priority'],
                    'plate_type': 'smaller_plates',
                    'plate': plate,
                    'num_plates': total_plates,
                    'num_cuts': 0,
                    'waste_percent': waste_percent,
                    'plates_across': plates_across
                }
            
            elif plate_type == 'larger':
                # Нужен разрез (ТОЛЬКО ВДОЛЬ плиты)
                if plate['length'] >= road_length:
                    plates_needed = 1  # одна плита разрезается
                    used_area = road_width * road_length
                    road_area = road_width * road_length
                    
                    waste_percent = ((used_area - road_area) / road_area * 100) if road_area > 0 else 0
                    
                    solution = {
                        'reinforcement': reinforcement,
                        'priority': plate['priority'],
                        'plate_type': 'cut_along_length',
                        'plate': plate,
                        'num_plates': plates_needed,
                        'num_cuts': 1,
                        'waste_percent': waste_percent
                    }
                else:
                    continue  # плита слишком короткая для разреза
            
            # Проверяем, лучше ли это решение
            if solution['waste_percent'] < min_waste:
                best_solution = solution
                min_waste = solution['waste_percent']
                print(f"  ✅ Лучшее решение: отходы {solution['waste_percent']:.2f}%, разрезы {solution['num_cuts']}")
    
    return best_solution

def distribute_by_days(solution, max_plates_per_day=50):
    """Распределяет плиты по дням"""
    if not solution:
        return []
    
    days = []
    current_day = 1
    plates_today = 0
    
    for i in range(solution['num_plates']):
        if plates_today >= max_plates_per_day:
            current_day += 1
            plates_today = 0
        
        plates_today += 1
    
    # Добавляем последний день
    if plates_today > 0:
        current_day += 1
    
    return current_day - 1  # возвращаем количество дней

def print_solution(solution, days):
    """Выводит решение"""
    if not solution:
        print("Решение не найдено")
        return
    
    print(f"\n{'='*60}")
    print(f"ОПТИМАЛЬНОЕ РЕШЕНИЕ")
    print(f"{'='*60}")
    print(f"Армирование: {solution['reinforcement']}")
    print(f"Приоритет срочности: {solution['priority']}")
    print(f"Тип решения: {solution['plate_type']}")
    print(f"Плита: {solution['plate']['marking']} ({solution['plate']['width']}x{solution['plate']['length']}м)")
    print(f"Количество плит: {solution['num_plates']}")
    print(f"Количество разрезов: {solution['num_cuts']}")
    print(f"Процент отходов: {solution['waste_percent']:.2f}%")
    
    if 'plates_across' in solution:
        print(f"Плит поперек: {solution['plates_across']}")
    
    print(f"\nДней работы: {days}")
    print(f"Плит в день: {math.ceil(solution['num_plates'] / days) if days > 0 else 0}")

def main():
    """Основная функция"""
    print("=== ОПТИМАЛЬНАЯ РАСКЛАДКА ПЛИТ ===")
    print("Критерии:")
    print("- Срочность по неделе армирования (чем меньше - тем срочнее)")
    print("- Плиты без недели армирования - в последнюю очередь")
    print("- Минимизация разрезов и остатков")
    print("- Одинаковое армирование в дорожке")
    print()
    
    # Загружаем и парсим плиты
    plates = load_and_parse_plates()
    print(f"Обработано {len(plates)} плит")
    
    if not plates:
        print("Не удалось загрузить плиты")
        return
    
    # Параметры дорожек
    road_width = 1.2
    road_length = 101
    max_plates_per_day = 50
    
    # Находим оптимальное решение
    solution = find_optimal_solution(plates, road_width, road_length)
    
    # Распределяем по дням
    days = distribute_by_days(solution, max_plates_per_day)
    
    # Выводим результат
    print_solution(solution, days)
    
    print(f"\n{'='*60}")
    print(f"ИТОГО: {days} дней работы")
    print(f"Всего плит: {solution['num_plates'] if solution else 0}")
    print(f"Отходы: {solution['waste_percent']:.2f}%" if solution else "0%")

if __name__ == "__main__":
    main()


