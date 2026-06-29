#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тесты модуля factory_cost

Проверяем:
- Чтение Excel
- Импорт в БД
- API для получения себестоимости
- Использование существующего парсера
"""

import os
import sys
import tempfile
import sqlite3

# Добавляем корень проекта в путь
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from factory_cost.db_schema import init_factory_cost_schema, get_factory_cost_stats, clear_factory_costs
from factory_cost.excel_reader import read_factory_costs_from_excel
from factory_cost.import_from_xlsx import import_factory_costs_from_xlsx
from factory_cost.cost_engine import (
    get_cost_by_plate_name,
    get_cost_by_params,
    get_all_available_plates,
    get_cost_breakdown
)


def test_db_schema():
    """Тест: Инициализация схемы БД"""
    print("\n[TEST] Инициализация схемы БД...")
    
    with tempfile.NamedTemporaryFile(suffix='.db', delete=False) as tmp:
        db_path = tmp.name
    
    try:
        # Инициализируем схему
        init_factory_cost_schema(db_path)
        
        # Проверяем наличие таблиц
        conn = sqlite3.connect(db_path)
        cur = conn.cursor()
        
        cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
        tables = [row[0] for row in cur.fetchall()]
        
        assert 'factory_plate_costs' in tables, "Таблица factory_plate_costs не создана"
        assert 'factory_plate_cost_components' in tables, "Таблица factory_plate_cost_components не создана"
        
        conn.close()
        
        print("✅ Схема БД инициализирована корректно")
        return True
        
    finally:
        if os.path.exists(db_path):
            os.remove(db_path)


def test_excel_reading():
    """Тест: Чтение Excel файла"""
    print("\n[TEST] Чтение Excel файла...")
    
    # Путь к тестовому файлу
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    xlsx_path = os.path.join(
        project_root,
        'банк знаний',
        'Расчет новых цен на ПБ 10.09.2025 (1).xlsx'
    )
    
    if not os.path.exists(xlsx_path):
        print("⚠️ Excel файл не найден, пропускаем тест")
        return True
    
    # Читаем
    data, kef = read_factory_costs_from_excel(xlsx_path)
    
    assert len(data) > 0, "Не прочитано ни одной записи"
    assert kef is not None, "КЭФ не найден"
    assert kef >= 1.0, f"КЭФ должен быть >= 1.0, получен {kef}"
    
    # Проверяем структуру первой записи
    first = data[0]
    required_keys = ['plate_name_excel', 'direct_cost', 'reinforcement_cost', 'concrete_cost']
    for key in required_keys:
        assert key in first, f"Отсутствует ключ {key} в данных"
    
    print(f"✅ Прочитано {len(data)} записей, КЭФ={kef}")
    return True


def test_import():
    """Тест: Импорт из Excel в БД"""
    print("\n[TEST] Импорт из Excel в БД...")
    
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    xlsx_path = os.path.join(
        project_root,
        'банк знаний',
        'Расчет новых цен на ПБ 10.09.2025 (1).xlsx'
    )
    
    if not os.path.exists(xlsx_path):
        print("⚠️ Excel файл не найден, пропускаем тест")
        return True
    
    with tempfile.NamedTemporaryFile(suffix='.db', delete=False) as tmp:
        db_path = tmp.name
    
    try:
        # Импортируем
        stats = import_factory_costs_from_xlsx(xlsx_path, db_path)
        
        assert stats['imported'] > 0, "Ничего не импортировано"
        assert stats['errors'] == 0, f"Есть ошибки при импорте: {stats['errors']}"
        
        # Проверяем статистику БД
        db_stats = get_factory_cost_stats(db_path)
        
        assert db_stats['total_plates'] > 0
        assert db_stats['total_plates'] <= stats['imported'], (
            "В БД не может быть больше уникальных плит, чем обработано строк Excel"
        )
        
        print(f"✅ Импортировано {stats['imported']} плит")
        return True
        
    finally:
        if os.path.exists(db_path):
            os.remove(db_path)


def test_cost_api():
    """Тест: API получения себестоимости"""
    print("\n[TEST] API получения себестоимости...")
    
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    xlsx_path = os.path.join(
        project_root,
        'банк знаний',
        'Расчет новых цен на ПБ 10.09.2025 (1).xlsx'
    )
    
    if not os.path.exists(xlsx_path):
        print("⚠️ Excel файл не найден, пропускаем тест")
        return True
    
    with tempfile.NamedTemporaryFile(suffix='.db', delete=False) as tmp:
        db_path = tmp.name
    
    try:
        # Импортируем данные
        import_factory_costs_from_xlsx(xlsx_path, db_path)
        
        # Получаем список всех плит
        all_plates = get_all_available_plates(db_path)
        assert len(all_plates) > 0, "Нет плит в БД"
        
        # Тестируем поиск по названию
        test_plate = all_plates[0]
        cost = get_cost_by_plate_name(test_plate['plate_name'], db_path)
        
        assert cost is not None, f"Не найдена себестоимость для {test_plate['plate_name']}"
        assert 'direct_cost' in cost, "Отсутствует direct_cost"
        assert 'full_cost_with_kef' in cost, "Отсутствует full_cost_with_kef"
        assert 'components' in cost, "Отсутствуют компоненты"
        
        # Тестируем поиск по параметрам
        length_m = test_plate['length_dm'] / 10.0
        width_m = test_plate['width_dm'] / 10.0
        
        cost2 = get_cost_by_params(length_m, width_m, db_path)
        assert cost2 is not None, f"Не найдена себестоимость для {length_m}м × {width_m}м"
        
        # Тестируем детализацию
        breakdown = get_cost_breakdown(test_plate['plate_name'], db_path)
        assert breakdown is not None, "Не получена детализация"
        assert 'breakdown' in breakdown, "Отсутствует breakdown"
        
        print(f"✅ API работает корректно")
        print(f"   Тестовая плита: {test_plate['plate_name']}")
        print(f"   Себестоимость: {cost['full_cost_with_kef']:.2f} руб")
        return True
        
    finally:
        if os.path.exists(db_path):
            os.remove(db_path)


def test_parser_integration():
    """Тест: Интеграция с существующим парсером"""
    print("\n[TEST] Интеграция с существующим парсером...")
    
    # Импортируем парсер
    from core.config_and_data import (
        parse_load_code_from_name,
        get_load_code_for_plate,
        make_plate_name,
    )
    
    # Тест 1: Парсинг нагрузки из названия
    test_cases = [
        ("Плиты ПБ 71-12-8п", 8),
        ("ПБ 69-12-10п", 10),
        ("ПБ 66-12-12,5п", 13),  # 12.5 округляется до 13
    ]
    
    for plate_name, expected_load in test_cases:
        load = parse_load_code_from_name(plate_name)
        assert load == expected_load, \
            f"Неверная нагрузка для {plate_name}: ожидалось {expected_load}, получено {load}"
    
    # Тест 2: Создание названия плиты
    plate_name = make_plate_name(7.1, 1.2, load_code=10)
    assert "71" in plate_name, "Длина не в названии"
    assert "12" in plate_name, "Ширина не в названии"
    assert "10п" in plate_name, "Нагрузка не в названии"
    
    print("✅ Интеграция с парсером работает")
    return True


def run_all_tests():
    """Запуск всех тестов"""
    print("="*70)
    print("ТЕСТЫ МОДУЛЯ factory_cost")
    print("="*70)
    
    tests = [
        ("Схема БД", test_db_schema),
        ("Чтение Excel", test_excel_reading),
        ("Импорт", test_import),
        ("Cost API", test_cost_api),
        ("Интеграция с парсером", test_parser_integration),
    ]
    
    passed = 0
    failed = 0
    
    for name, test_func in tests:
        try:
            if test_func():
                passed += 1
            else:
                failed += 1
                print(f"❌ Тест '{name}' провален")
        except Exception as e:
            failed += 1
            print(f"❌ Тест '{name}' вызвал исключение: {e}")
            import traceback
            traceback.print_exc()
    
    print("\n" + "="*70)
    print(f"РЕЗУЛЬТАТ: {passed} пройдено, {failed} провалено")
    print("="*70)
    
    return failed == 0


if __name__ == '__main__':
    success = run_all_tests()
    sys.exit(0 if success else 1)

