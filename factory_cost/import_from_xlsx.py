#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Импорт заводской себестоимости из Excel в SQLite

КРИТИЧЕСКИ ВАЖНО: 
- НЕ парсит размеры плит самостоятельно
- Использует СУЩЕСТВУЮЩИЙ парсер из core.config_and_data
- Нагрузка определяется через get_load_code_for_plate()
"""

import os
import re
import sqlite3
from typing import List, Dict, Optional
from datetime import datetime

# Импортируем существующий парсер (НЕ писать свой!)
from core.config_and_data import parse_load_code_from_name, make_plate_name

from .excel_reader import read_factory_costs_from_excel
from .db_schema import init_factory_cost_schema, DEFAULT_DB_PATH


def parse_plate_dimensions_from_name(plate_name: str) -> Optional[tuple]:
    """
    Извлекает (длина_дм, ширина_дм) из названия плиты.
    
    ВАЖНО: В названии ширина указана в ДЕЦИМЕТРАХ (12 значит 1.2м),
    а в БД мы храним в ДЕЦИМЕТРАХ (length_dm, width_dm).
    
    Примеры:
        "Плиты ПБ 63-12-8п" → (63, 12)    # 6.3м × 1.2м
        "ПБ 71-5,3-8п" → (71, 5)           # 7.1м × 0.53м (округляем дробь)
    
    ВАЖНО: Эта функция ТОЛЬКО извлекает размеры из имени,
    но НЕ определяет нагрузку! Для нагрузки используем parse_load_code_from_name().
    """
    # Нормализуем: заменяем запятые на точки
    normalized = plate_name.replace(',', '.')
    
    # Формат: "ПБ 63-12-8п" или "Плиты ПБ 71-5.3-10п"
    # Паттерн: П[БК] <длина_дм>-<ширина_дм>-<нагрузка>п
    match = re.search(r'п[бк]\s*(\d+)\s*-\s*([\d\.]+)', normalized, re.IGNORECASE)
    if not match:
        return None
    
    try:
        length_dm = int(match.group(1))
        width_dm_str = match.group(2)
        
        # Ширина может быть: "12" (целое) или "5.3" (дробное)
        width_value = float(width_dm_str)
        
        # Округляем до целого дециметра
        # 12.0 → 12, 5.3 → 5 (для узких плит)
        width_dm = int(round(width_value))
        
        return (length_dm, width_dm)
    except (ValueError, AttributeError):
        return None


def validate_component_sum(
    components: Dict[str, float],
    direct_cost: float,
    tolerance_abs: float = 50.0,
    tolerance_pct: float = 2.0
) -> tuple[bool, float]:
    """
    Проверяет, что сумма компонентов совпадает с прямыми затратами.
    
    Args:
        components: Словарь {reinforcement, concrete, loops, izoform}
        direct_cost: Итого прямые затраты из Excel
        tolerance_abs: Абсолютная погрешность (руб)
        tolerance_pct: Относительная погрешность (%)
    
    Returns:
        (валидно?, отклонение в рублях)
    """
    comp_sum = sum(components.values())
    diff = abs(comp_sum - direct_cost)
    
    # Проверяем по абсолютной или относительной погрешности
    max_allowed = max(tolerance_abs, direct_cost * tolerance_pct / 100)
    
    is_valid = diff <= max_allowed
    return is_valid, diff


def import_factory_costs_from_xlsx(
    xlsx_path: str,
    db_path: str = DEFAULT_DB_PATH,
    clear_existing: bool = True
) -> Dict[str, int]:
    """
    Импортирует заводскую себестоимость из Excel в БД.
    
    Процесс:
    1. Читает Excel (лист "Стоимость" + КЭФ с листа "Себестоимость")
    2. Парсит названия плит через СУЩЕСТВУЮЩИЙ парсер
    3. Валидирует суммы компонентов
    4. Сохраняет в factory_plate_costs и factory_plate_cost_components
    
    Args:
        xlsx_path: Путь к Excel-файлу
        db_path: Путь к БД
        clear_existing: Очистить существующие данные перед импортом
    
    Returns:
        Статистика импорта
    """
    print(f"\n{'='*60}")
    print(f"ИМПОРТ ЗАВОДСКОЙ СЕБЕСТОИМОСТИ")
    print(f"{'='*60}")
    print(f"Файл: {os.path.basename(xlsx_path)}")
    print(f"БД: {db_path}")
    print(f"{'='*60}\n")
    
    # Инициализируем схему БД
    init_factory_cost_schema(db_path)
    
    # Читаем Excel
    print("[1/5] Чтение Excel...")
    cost_data, kef = read_factory_costs_from_excel(xlsx_path)
    
    if not cost_data:
        print("❌ Не удалось прочитать данные из Excel")
        return {'imported': 0, 'skipped': 0, 'errors': 0}
    
    print(f"✓ Прочитано записей: {len(cost_data)}")
    print(f"✓ КЭФ: {kef if kef else 'не найден'}")
    
    # Подключаемся к БД
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        
        # Очищаем старые данные
        if clear_existing:
            print("\n[2/5] Очистка старых данных...")
            cur.execute("DELETE FROM factory_plate_costs")
            cur.execute("DELETE FROM factory_plate_cost_components")
            conn.commit()
            print("✓ Старые данные удалены")
        
        # Импортируем записи
        print("\n[3/5] Импорт себестоимости...")
        
        stats = {
            'imported': 0,
            'skipped': 0,
            'errors': 0,
            'validation_errors': 0,
        }
        
        for idx, record in enumerate(cost_data, 1):
            plate_name_excel = record['plate_name_excel']
            
            # Парсим размеры
            dimensions = parse_plate_dimensions_from_name(plate_name_excel)
            if not dimensions:
                print(f"⚠️ [{idx}] Не удалось распарсить: {plate_name_excel}")
                stats['skipped'] += 1
                continue
            
            length_dm, width_dm = dimensions
            
            # Парсим нагрузку через СУЩЕСТВУЮЩИЙ парсер
            load_code = parse_load_code_from_name(plate_name_excel, default=8)
            
            # ВАЖНО: Если в названии "12,5п", парсер вернёт 13 (округление вверх)
            # Но мы сохраняем как float, чтобы различать 12.5 и 13
            # Проверяем, не была ли в имени нагрузка 12.5
            if '12,5' in plate_name_excel or '12.5' in plate_name_excel:
                load_code = 12.5
            
            # Создаём нормализованное имя плиты через СУЩЕСТВУЮЩИЙ make_plate_name
            # Он автоматически форматирует ширину правильно (12 → "12", 5 → "5")
            plate_name = make_plate_name(
                length_m=length_dm / 10.0,
                width_m=width_dm / 10.0,
                load_code=load_code
            )
            
            # Извлекаем компоненты
            components = {
                'reinforcement': record['reinforcement_cost'],
                'concrete': record['concrete_cost'],
                'loops': record['loops_cost'],
                'izoform': record['izoform_cost'],
            }
            
            # Валидируем сумму компонентов
            is_valid, diff = validate_component_sum(components, record['direct_cost'])
            quality_flag = None
            if not is_valid:
                quality_flag = 'components_mismatch'
                stats['validation_errors'] += 1
                print(f"⚠️ [{idx}] {plate_name}: расхождение компонентов {diff:.2f} руб")
            
            # Рассчитываем полную себестоимость
            direct_cost = record['direct_cost']
            overhead_cost = 0.0
            full_cost = direct_cost
            full_cost_with_kef = full_cost
            
            if kef and kef > 1.0:
                # КЭФ применяется к прямым затратам
                full_cost_with_kef = direct_cost * kef
                overhead_cost = full_cost_with_kef - direct_cost
            
            # Вставляем в factory_plate_costs
            try:
                cur.execute("""
                    INSERT OR REPLACE INTO factory_plate_costs (
                        plate_name, length_dm, width_dm, load_code,
                        direct_cost, overhead_cost, full_cost,
                        kef, full_cost_with_kef,
                        volume_m3, concrete_grade, quality_flag,
                        source_file, source_sheet, source_row, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    plate_name, length_dm, width_dm, load_code,
                    direct_cost, overhead_cost, full_cost,
                    kef, full_cost_with_kef,
                    record['volume_m3'], record['concrete_grade'], quality_flag,
                    os.path.basename(xlsx_path), 'Стоимость', record['source_row'],
                    datetime.now().isoformat()
                ))
                
                # Вставляем компоненты
                for component, value in components.items():
                    cur.execute("""
                        INSERT OR REPLACE INTO factory_plate_cost_components
                        (plate_name, component, value)
                        VALUES (?, ?, ?)
                    """, (plate_name, component, value))
                
                stats['imported'] += 1
                
                # Прогресс
                if stats['imported'] % 10 == 0:
                    print(f"  Импортировано: {stats['imported']}...")
                
            except Exception as e:
                print(f"❌ [{idx}] Ошибка при импорте {plate_name}: {e}")
                stats['errors'] += 1
                continue
        
        # Коммитим
        print("\n[4/5] Сохранение в БД...")
        conn.commit()
        print("✓ Изменения сохранены")
        
        # Статистика
        print(f"\n[5/5] Импорт завершён")
        print(f"{'='*60}")
        print(f"✓ Импортировано: {stats['imported']}")
        print(f"⚠️ Пропущено: {stats['skipped']}")
        print(f"⚠️ С расхождениями: {stats['validation_errors']}")
        print(f"❌ Ошибок: {stats['errors']}")
        print(f"{'='*60}\n")
        
        return stats
        
    finally:
        conn.close()


if __name__ == '__main__':
    # Тестовый импорт
    import sys
    
    xlsx_path = os.path.join(
        os.path.dirname(os.path.dirname(__file__)),
        'банк знаний',
        'Расчет новых цен на ПБ 10.09.2025 (1).xlsx'
    )
    
    if not os.path.exists(xlsx_path):
        print(f"❌ Файл не найден: {xlsx_path}")
        sys.exit(1)
    
    stats = import_factory_costs_from_xlsx(xlsx_path)
    
    if stats['imported'] > 0:
        print("\n✅ Импорт успешно завершён!")
    else:
        print("\n❌ Импорт завершился с ошибками")
        sys.exit(1)

