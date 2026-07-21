#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт валидации заводской себестоимости в БД

Проверяет:
- Корректность сумм компонентов
- Адекватность значений (нет отрицательных, нет слишком больших)
- Дубликаты
- Плиты без компонентов

Использование:
    python scripts/validate_factory_costs.py
    python scripts/validate_factory_costs.py --detailed  # подробный отчёт
"""

import os
import sys
import argparse
import sqlite3

# Добавляем корень проекта в путь
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from factory_cost.db_schema import DEFAULT_DB_PATH, get_factory_cost_stats


def validate_component_sums(db_path: str) -> list:
    """
    Проверяет, что сумма компонентов совпадает с прямыми затратами.
    
    Returns:
        Список плит с расхождениями
    """
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        
        # Получаем плиты с компонентами
        cur.execute("""
            SELECT 
                c.plate_name,
                c.direct_cost,
                COALESCE(SUM(comp.value), 0) as components_sum
            FROM factory_plate_costs c
            LEFT JOIN factory_plate_cost_components comp ON c.plate_name = comp.plate_name
            GROUP BY c.plate_name, c.direct_cost
        """)
        
        problems = []
        for row in cur.fetchall():
            plate_name, direct_cost, comp_sum = row
            diff = abs(direct_cost - comp_sum)
            
            # Допуск: max(50 руб, 2%)
            tolerance = max(50, direct_cost * 0.02)
            
            if diff > tolerance:
                problems.append({
                    'plate_name': plate_name,
                    'direct_cost': direct_cost,
                    'components_sum': comp_sum,
                    'difference': diff,
                })
        
        return problems
        
    finally:
        conn.close()


def validate_reasonable_values(db_path: str) -> list:
    """
    Проверяет адекватность значений (нет отрицательных, нет слишком больших).
    
    Returns:
        Список плит с неадекватными значениями
    """
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        
        cur.execute("""
            SELECT 
                plate_name, length_dm, width_dm, load_code,
                direct_cost, full_cost_with_kef, volume_m3
            FROM factory_plate_costs
        """)
        
        problems = []
        for row in cur.fetchall():
            plate_name, length_dm, width_dm, load_code, direct, full, volume = row
            
            issues = []
            
            # Проверка отрицательных значений
            if direct < 0:
                issues.append(f"отрицательная прямая себестоимость: {direct}")
            if full < 0:
                issues.append(f"отрицательная полная себестоимость: {full}")
            if volume and volume < 0:
                issues.append(f"отрицательный объём: {volume}")
            
            # Проверка слишком больших значений
            if direct > 100000:
                issues.append(f"подозрительно высокая себестоимость: {direct}")
            
            # Проверка объёма (не должен быть больше 2 м³ для плит)
            if volume and volume > 2.0:
                issues.append(f"подозрительно большой объём: {volume} м³")
            
            # Проверка размеров
            if length_dm > 300:  # 30 метров - явно ошибка
                issues.append(f"подозрительная длина: {length_dm} дм")
            if width_dm > 200:  # 20 метров - явно ошибка
                issues.append(f"подозрительная ширина: {width_dm} дм")
            
            # Проверка нагрузки
            if load_code < 4 or load_code > 20:
                issues.append(f"подозрительная нагрузка: {load_code}")
            
            if issues:
                problems.append({
                    'plate_name': plate_name,
                    'issues': issues,
                })
        
        return problems
        
    finally:
        conn.close()


def find_duplicates(db_path: str) -> list:
    """
    Находит дубликаты по (длина, ширина, нагрузка).
    
    Returns:
        Список дубликатов
    """
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        
        cur.execute("""
            SELECT length_dm, width_dm, load_code, COUNT(*) as cnt
            FROM factory_plate_costs
            GROUP BY length_dm, width_dm, load_code
            HAVING cnt > 1
        """)
        
        duplicates = []
        for row in cur.fetchall():
            length_dm, width_dm, load_code, cnt = row
            duplicates.append({
                'dimensions': f"{length_dm}дм × {width_dm}дм × {load_code}п",
                'count': cnt,
            })
        
        return duplicates
        
    finally:
        conn.close()


def find_plates_without_components(db_path: str) -> list:
    """
    Находит плиты без компонентов.
    
    Returns:
        Список плит без компонентов
    """
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        
        cur.execute("""
            SELECT c.plate_name
            FROM factory_plate_costs c
            LEFT JOIN factory_plate_cost_components comp ON c.plate_name = comp.plate_name
            WHERE comp.plate_name IS NULL
        """)
        
        return [row[0] for row in cur.fetchall()]
        
    finally:
        conn.close()


def main():
    parser = argparse.ArgumentParser(
        description='Валидация заводской себестоимости в БД'
    )
    parser.add_argument(
        '--detailed',
        action='store_true',
        help='Показать подробный отчёт с деталями'
    )
    parser.add_argument(
        '--db',
        type=str,
        default=None,
        help='Путь к БД (по умолчанию: pb.db в корне проекта)'
    )
    
    args = parser.parse_args()
    
    # Определяем путь к БД
    if args.db:
        db_path = args.db
    else:
        db_path = DEFAULT_DB_PATH
    
    if not os.path.exists(db_path):
        print(f"❌ БД не найдена: {db_path}")
        sys.exit(1)
    
    print(f"\n{'='*70}")
    print(f"ВАЛИДАЦИЯ ЗАВОДСКОЙ СЕБЕСТОИМОСТИ")
    print(f"{'='*70}")
    print(f"БД: {db_path}")
    print(f"{'='*70}\n")
    
    # Статистика
    try:
        stats = get_factory_cost_stats(db_path)
        print(f"📊 Статистика:")
        print(f"   Всего плит: {stats['total_plates']}")
        print(f"   С проблемами (флаг): {stats['problem_plates']}")
        print()
    except Exception as e:
        print(f"⚠️ Не удалось получить статистику: {e}\n")
    
    # Запуск валидации
    all_ok = True
    
    # 1. Проверка сумм компонентов
    print("[1/4] Проверка сумм компонентов...")
    component_problems = validate_component_sums(db_path)
    if component_problems:
        all_ok = False
        print(f"❌ Найдено плит с расхождениями: {len(component_problems)}")
        if args.detailed:
            for p in component_problems[:10]:  # Показываем первые 10
                print(f"   - {p['plate_name']}: "
                      f"прямые={p['direct_cost']:.2f}, "
                      f"компоненты={p['components_sum']:.2f}, "
                      f"разница={p['difference']:.2f}")
            if len(component_problems) > 10:
                print(f"   ... и ещё {len(component_problems) - 10}")
    else:
        print("✅ Все суммы компонентов корректны")
    print()
    
    # 2. Проверка адекватности значений
    print("[2/4] Проверка адекватности значений...")
    value_problems = validate_reasonable_values(db_path)
    if value_problems:
        all_ok = False
        print(f"❌ Найдено плит с подозрительными значениями: {len(value_problems)}")
        if args.detailed:
            for p in value_problems[:10]:
                print(f"   - {p['plate_name']}:")
                for issue in p['issues']:
                    print(f"     • {issue}")
            if len(value_problems) > 10:
                print(f"   ... и ещё {len(value_problems) - 10}")
    else:
        print("✅ Все значения в разумных пределах")
    print()
    
    # 3. Поиск дубликатов
    print("[3/4] Поиск дубликатов...")
    duplicates = find_duplicates(db_path)
    if duplicates:
        all_ok = False
        print(f"❌ Найдено дубликатов: {len(duplicates)}")
        if args.detailed:
            for d in duplicates:
                print(f"   - {d['dimensions']}: {d['count']} записей")
    else:
        print("✅ Дубликатов не найдено")
    print()
    
    # 4. Плиты без компонентов
    print("[4/4] Поиск плит без компонентов...")
    no_components = find_plates_without_components(db_path)
    if no_components:
        all_ok = False
        print(f"❌ Найдено плит без компонентов: {len(no_components)}")
        if args.detailed:
            for plate in no_components[:10]:
                print(f"   - {plate}")
            if len(no_components) > 10:
                print(f"   ... и ещё {len(no_components) - 10}")
    else:
        print("✅ У всех плит есть компоненты")
    print()
    
    # Итог
    print(f"{'='*70}")
    if all_ok:
        print("✅ ВАЛИДАЦИЯ ПРОЙДЕНА: Проблем не обнаружено")
        sys.exit(0)
    else:
        print("⚠️ ВАЛИДАЦИЯ ЗАВЕРШЕНА С ПРЕДУПРЕЖДЕНИЯМИ")
        print("\nИспользуйте флаг --detailed для подробного отчёта")
        sys.exit(1)


if __name__ == '__main__':
    main()

