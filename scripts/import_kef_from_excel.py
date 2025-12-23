#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для импорта значений КЭФ из Excel в БД

Использование:
    python scripts/import_kef_from_excel.py
    python scripts/import_kef_from_excel.py --file "путь/к/файлу.xlsx"
"""

import os
import sys

# Добавляем корневую директорию в путь
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE_DIR)

from cost_calculation.load_from_excel import load_kef_from_excel
import core.config_and_data as cfg


def main():
    """Основная функция импорта КЭФ"""
    import argparse
    
    parser = argparse.ArgumentParser(description='Импорт значений КЭФ из Excel в БД')
    parser.add_argument(
        '--file',
        type=str,
        help='Путь к Excel файлу (по умолчанию используется файл из банк знаний)'
    )
    
    args = parser.parse_args()
    
    print("=" * 80)
    print("ИМПОРТ ЗНАЧЕНИЙ КЭФ ИЗ EXCEL")
    print("=" * 80)
    print()
    
    db_path = cfg.PRICE_DB_PATH
    
    if args.file:
        # Если указан файл, временно меняем путь
        original_path = cfg.BASE_DIR
        excel_path = args.file
        print(f"📁 Используется файл: {excel_path}")
        
        # Временно меняем EXCEL_PATH в модуле
        import cost_calculation.load_from_excel as load_module
        load_module.EXCEL_PATH = excel_path
    
    print(f"📊 БД: {db_path}")
    print()
    
    try:
        load_kef_from_excel(db_path)
        print()
        print("=" * 80)
        print("✅ ИМПОРТ ЗАВЕРШЕН")
        print("=" * 80)
    except Exception as e:
        print(f"\n❌ Ошибка при импорте: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()

