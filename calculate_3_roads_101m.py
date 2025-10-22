#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Расчет оптимальной раскладки плит для 3 дорожек по 101 метр
"""

import sqlite3
import pandas as pd
import math
import re
from collections import defaultdict

def load_plates_from_db():
    """Загружает плиты из базы данных"""
    conn = sqlite3.connect('pb.db')
    df = pd.read_sql_query("SELECT * FROM plity_ex", conn)
    conn.close()
    return df

def parse_plate_data(df):
    """Парсит данные плит из маркировки"""
    plates = []
    
    for idx, row in df.iterrows():
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

def find_optimal_solution_for_roads(plates, road_width=1.2, road_length=101, num_roads=3):
    """Находит оптимальное решение для дорожек"""
    
    total_road_length = road_length * num_roads
    
    print(f"=== ПОИСК ОПТИМАЛЬНОГО РЕШЕНИЯ ===")
    print(f"Ширина дорожки: {road_width}м")
    print(f"Длина дорожки: {road_length}м")
    print(f"Количество дорожек: {num_roads}")
    print(f"Общая длина: {total_road_length}м")
    print()
    
    # Сортируем плиты по приоритету (срочности)
    sorted_plates = sorted(plates, key=lambda x: x['priority'])
    
    print("Плиты отсортированы по срочности:")
    for i, plate in enumerate(sorted_plates[:10]):
        print(f"  {i+1}. Приоритет {plate['priority']}: {plate['marking']} ({plate['width']}x{plate['length']}м)")
    
    # Группируем по армированию
    plates_by_reinforcement = defaultdict(list)
    for plate in sorted_plates:
        plates_by_reinforcement[plate['reinforcement']].append(plate)
    
    best_solution = None
    min_waste = float('inf')
    min_cuts = float('inf')
    
    print(f"\nПроверяем комбинации по армированию:")
    
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
                plates_needed = math.ceil(total_road_length / plate['length'])
                total_length = plates_needed * plate['length']
                
                used_area = plates_needed * plate['width'] * plate['length']
                road_area = road_width * total_road_length
                
                waste_percent = ((used_area - road_area) / road_area * 100) if road_area > 0 else 0
                
                solution = {
                    'reinforcement': reinforcement,
                    'priority': plate['priority'],
                    'plate_type': 'exact_width',
                    'plate': plate,
                    'num_plates': plates_needed,
                    'num_cuts': 0,
                    'waste_percent': waste_percent,
                    'total_length': total_length,
                    'plates_per_road': plates_needed // num_roads
                }
                
            elif plate_type == 'smaller':
                # Несколько плит поперек
                plates_across = math.ceil(road_width / plate['width'])
                plates_length = math.ceil(total_road_length / plate['length'])
                total_plates = plates_across * plates_length
                
                used_area = total_plates * plate['width'] * plate['length']
                road_area = road_width * total_road_length
                
                waste_percent = ((used_area - road_area) / road_area * 100) if road_area > 0 else 0
                
                solution = {
                    'reinforcement': reinforcement,
                    'priority': plate['priority'],
                    'plate_type': 'smaller_plates',
                    'plate': plate,
                    'num_plates': total_plates,
                    'num_cuts': 0,
                    'waste_percent': waste_percent,
                    'plates_across': plates_across,
                    'plates_per_road': total_plates // num_roads
                }
            
            elif plate_type == 'larger':
                # Нужен разрез (ТОЛЬКО ВДОЛЬ плиты)
                if plate['length'] >= road_length:
                    plates_needed = num_roads  # по одной плите на дорожку
                    used_area = road_width * total_road_length
                    road_area = road_width * total_road_length
                    
                    waste_percent = ((used_area - road_area) / road_area * 100) if road_area > 0 else 0
                    
                    solution = {
                        'reinforcement': reinforcement,
                        'priority': plate['priority'],
                        'plate_type': 'cut_along_length',
                        'plate': plate,
                        'num_plates': plates_needed,
                        'num_cuts': plates_needed,
                        'waste_percent': waste_percent,
                        'plates_per_road': 1
                    }
                else:
                    continue  # плита слишком короткая для разреза
            
            # Проверяем, лучше ли это решение
            is_better = False
            if solution['waste_percent'] < min_waste:
                is_better = True
            elif solution['waste_percent'] == min_waste and solution['num_cuts'] < min_cuts:
                is_better = True
            elif (solution['waste_percent'] == min_waste and 
                  solution['num_cuts'] == min_cuts and 
                  solution['priority'] < best_solution['priority'] if best_solution else True):
                is_better = True
            
            if is_better:
                best_solution = solution
                min_waste = solution['waste_percent']
                min_cuts = solution['num_cuts']
                print(f"  ✅ Лучшее решение: отходы {solution['waste_percent']:.2f}%, разрезы {solution['num_cuts']}")
    
    return best_solution

def distribute_by_days(solution, max_plates_per_day=50):
    """Распределяет плиты по дням"""
    if not solution:
        return []
    
    days_needed = math.ceil(solution['num_plates'] / max_plates_per_day)
    
    days = []
    current_day = 1
    plates_today = 0
    
    for i in range(solution['num_plates']):
        if plates_today >= max_plates_per_day:
            current_day += 1
            plates_today = 0
        
        plates_today += 1
    
    return current_day

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
    print(f"Плита: {solution['plate']['marking']}")
    print(f"Размер плиты: {solution['plate']['width']}x{solution['plate']['length']}м")
    print(f"Количество плит: {solution['num_plates']}")
    print(f"Плит на дорожку: {solution['plates_per_road']}")
    print(f"Количество разрезов: {solution['num_cuts']}")
    print(f"Процент отходов: {solution['waste_percent']:.2f}%")
    
    if 'plates_across' in solution:
        print(f"Плит поперек дорожки: {solution['plates_across']}")
    
    print(f"\nРАСПРЕДЕЛЕНИЕ ПО ДНЯМ:")
    print(f"Дней работы: {days}")
    print(f"Плит в день: {math.ceil(solution['num_plates'] / days) if days > 0 else 0}")
    
    print(f"\nДЕТАЛЬНАЯ РАСКЛАДКА:")
    print(f"Дорожка 1: {solution['plates_per_road']} плит {solution['plate']['marking']}")
    print(f"Дорожка 2: {solution['plates_per_road']} плит {solution['plate']['marking']}")
    print(f"Дорожка 3: {solution['plates_per_road']} плит {solution['plate']['marking']}")

def main():
    """Основная функция"""
    print("=== РАСЧЕТ ОПТИМАЛЬНОЙ РАСКЛАДКИ ПЛИТ ===")
    print("Задача: 3 дорожки по 101 метр каждая")
    print("Критерии:")
    print("- Срочность по неделе армирования (чем меньше - тем срочнее)")
    print("- Плиты без недели армирования - в последнюю очередь")
    print("- Минимизация разрезов и остатков")
    print("- Одинаковое армирование в дорожке")
    print()
    
    # Загружаем данные
    df = load_plates_from_db()
    print(f"Загружено {len(df)} записей из базы данных")
    
    # Парсим плиты
    plates = parse_plate_data(df)
    print(f"Обработано {len(plates)} плит")
    
    if not plates:
        print("Не удалось загрузить плиты")
        return
    
    # Параметры дорожек
    road_width = 1.2
    road_length = 101
    num_roads = 3
    max_plates_per_day = 50
    
    # Находим оптимальное решение
    solution = find_optimal_solution_for_roads(plates, road_width, road_length, num_roads)
    
    # Распределяем по дням
    days = distribute_by_days(solution, max_plates_per_day)
    
    # Выводим результат
    print_solution(solution, days)
    
    print(f"\n{'='*60}")
    print(f"ИТОГО:")
    print(f"- Дней работы: {days}")
    print(f"- Всего плит: {solution['num_plates'] if solution else 0}")
    print(f"- Отходы: {solution['waste_percent']:.2f}%" if solution else "0%")
    print(f"- Разрезы: {solution['num_cuts'] if solution else 0}")

if __name__ == "__main__":
    main()


