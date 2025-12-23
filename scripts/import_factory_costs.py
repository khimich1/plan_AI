#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт импорта заводской себестоимости из Excel в БД

Использование:
    python scripts/import_factory_costs.py
    python scripts/import_factory_costs.py --file "путь/к/файлу.xlsx"
    python scripts/import_factory_costs.py --no-clear  # не очищать старые данные
"""

import os
import sys
import argparse

# Добавляем корень проекта в путь
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from factory_cost.import_from_xlsx import import_factory_costs_from_xlsx
from factory_cost.db_schema import get_factory_cost_stats


def main():
    parser = argparse.ArgumentParser(
        description='Импорт заводской себестоимости из Excel в БД'
    )
    parser.add_argument(
        '--file',
        type=str,
        default=None,
        help='Путь к Excel-файлу (по умолчанию: банк знаний/Расчет новых цен на ПБ 10.09.2025 (1).xlsx)'
    )
    parser.add_argument(
        '--no-clear',
        action='store_true',
        help='Не очищать существующие данные перед импортом'
    )
    parser.add_argument(
        '--db',
        type=str,
        default=None,
        help='Путь к БД (по умолчанию: pb.db в корне проекта)'
    )
    
    args = parser.parse_args()
    
    # Определяем путь к Excel
    if args.file:
        xlsx_path = args.file
    else:
        # Дефолтный путь
        project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        xlsx_path = os.path.join(
            project_root,
            'банк знаний',
            'Расчет новых цен на ПБ 10.09.2025 (1).xlsx'
        )
    
    # Проверяем существование файла
    if not os.path.exists(xlsx_path):
        print(f"❌ Файл не найден: {xlsx_path}")
        print("\nИспользуйте --file для указания пути к файлу")
        sys.exit(1)
    
    # Определяем путь к БД
    if args.db:
        db_path = args.db
    else:
        project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        db_path = os.path.join(project_root, 'pb.db')
    
    print(f"\n{'='*70}")
    print(f"СКРИПТ ИМПОРТА ЗАВОДСКОЙ СЕБЕСТОИМОСТИ")
    print(f"{'='*70}")
    print(f"Excel: {os.path.basename(xlsx_path)}")
    print(f"БД: {db_path}")
    print(f"Режим: {'добавление' if args.no_clear else 'перезапись'}")
    print(f"{'='*70}\n")
    
    # Показываем текущее состояние БД
    try:
        stats_before = get_factory_cost_stats(db_path)
        print(f"[*] Состояние БД ДО импорта:")
        print(f"   Плит в БД: {stats_before['total_plates']}")
        if stats_before['total_plates'] > 0:
            print(f"   С проблемами: {stats_before['problem_plates']}")
            print(f"   Нагрузки: {stats_before['load_codes']}")
        print()
    except Exception as e:
        print(f"[!] Не удалось получить статистику до импорта: {e}\n")
    
    # Запуск импорта
    try:
        stats = import_factory_costs_from_xlsx(
            xlsx_path=xlsx_path,
            db_path=db_path,
            clear_existing=not args.no_clear
        )
    except Exception as e:
        print(f"\n❌ ОШИБКА при импорте: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
    
    # Показываем результат
    print(f"\n{'='*70}")
    print(f"РЕЗУЛЬТАТ ИМПОРТА")
    print(f"{'='*70}")
    print(f"✅ Успешно импортировано: {stats['imported']}")
    if stats['skipped'] > 0:
        print(f"⚠️  Пропущено: {stats['skipped']}")
    if stats.get('validation_errors', 0) > 0:
        print(f"⚠️  С расхождением компонентов: {stats['validation_errors']}")
    if stats['errors'] > 0:
        print(f"❌ Ошибок: {stats['errors']}")
    print(f"{'='*70}\n")
    
    # Показываем итоговую статистику БД
    try:
        stats_after = get_factory_cost_stats(db_path)
        print(f"[*] Состояние БД ПОСЛЕ импорта:")
        print(f"   Всего плит: {stats_after['total_plates']}")
        print(f"   С проблемами: {stats_after['problem_plates']}")
        print(f"   Длины (дм): {stats_after['length_range_dm'][0]} - {stats_after['length_range_dm'][1]}")
        print(f"   Ширины (дм): {stats_after['width_range_dm'][0]} - {stats_after['width_range_dm'][1]}")
        print(f"   Нагрузки: {stats_after['load_codes']}")
        print()
    except Exception as e:
        print(f"[!] Не удалось получить статистику после импорта: {e}\n")
    
    # Завершение
    if stats['imported'] > 0:
        print("✅ Импорт успешно завершён!")
        sys.exit(0)
    else:
        print("❌ Ничего не импортировано")
        sys.exit(1)


if __name__ == '__main__':
    main()

