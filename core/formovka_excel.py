#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль создания Excel-файлов формовки по дорожкам.
Заполняет шаблон "!КЗ ПБ Шаблон.xlsx" данными о плитах на дорожке.
"""
import os
import shutil
from datetime import datetime
from typing import List, Dict, Optional

try:
    from openpyxl import load_workbook
    from openpyxl.styles import Font
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False
    print("[FORMOVKA] ⚠️ openpyxl не установлен. Установите: pip install openpyxl")


def create_formovka_excel(
    track_number: int,
    max_reinforcement: float,
    plates_info: List[Dict],
    output_dir: str,
    template_path: str = "банк знаний/!КЗ ПБ Шаблон.xlsx",
    date_str: Optional[str] = None
) -> Optional[str]:
    """
    Создает Excel-файл формовки для дорожки по шаблону.
    
    Args:
        track_number: Номер дорожки
        max_reinforcement: Максимальное армирование на дорожке
        plates_info: Список словарей с информацией о плитах:
            - plate_name: название плиты (например "ПБ 80-12-10п")
            - qty: количество штук
            - kp_date: срок/дата заказа (опционально)
            - customer: заказчик (опционально)
        output_dir: Директория для сохранения файла
        template_path: Путь к шаблону Excel
        date_str: Дата формовки (по умолчанию текущая)
    
    Returns:
        Путь к созданному файлу или None в случае ошибки
    """
    if not OPENPYXL_AVAILABLE:
        print("[FORMOVKA] ❌ openpyxl не установлен")
        return None
    
    if not os.path.exists(template_path):
        print(f"[FORMOVKA] ❌ Шаблон не найден: {template_path}")
        return None
    
    if not plates_info:
        print(f"[FORMOVKA] ⚠️ Нет данных о плитах для дорожки {track_number}")
        return None
    
    try:
        # Создаем директорию если не существует
        os.makedirs(output_dir, exist_ok=True)
        
        # Генерируем имя выходного файла
        if date_str is None:
            date_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"Формовка_Дорожка_{track_number}_{date_str}.xlsx"
        output_path = os.path.join(output_dir, output_filename)
        
        # Копируем шаблон
        shutil.copy2(template_path, output_path)
        
        # Открываем скопированный файл
        wb = load_workbook(output_path)
        ws = wb.active  # Первый лист
        
        # ========== ЗАПОЛНЯЕМ ДАННЫЕ ==========
        
        # Формат даты
        current_date = datetime.now().strftime("%d.%m.%Y")
        
        # Функция для безопасной записи в ячейку (может быть merged)
        def safe_set_cell(cell_ref, value):
            try:
                cell = ws[cell_ref]
                # Проверяем, не является ли ячейка частью объединённого диапазона
                if hasattr(cell, 'value'):
                    cell.value = value
                    return True
            except Exception as e:
                print(f"[FORMOVKA] Предупреждение: не удалось записать в {cell_ref}: {e}")
                # Пробуем использовать cell напрямую
                try:
                    ws.cell(row=cell.row, column=cell.column).value = value
                    return True
                except:
                    pass
            return False
        
        # 1. Дата формовки (ячейка B3, согласно скриншоту)
        safe_set_cell('B3', current_date)
        # Устанавливаем размер шрифта 32 для B3
        try:
            ws['B3'].font = Font(size=32)
        except Exception as e:
            print(f"[FORMOVKA] ⚠️ Не удалось установить шрифт для B3: {e}")
        
        # 2. Дата приемки ОТК (ячейка B7)
        safe_set_cell('B7', current_date)
        # Устанавливаем размер шрифта 32 для B7
        try:
            ws['B7'].font = Font(size=32)
        except Exception as e:
            print(f"[FORMOVKA] ⚠️ Не удалось установить шрифт для B7: {e}")
        
        # 3. Номер дорожки (ячейка D3 - "№ дорожки для формовки")
        safe_set_cell('D3', track_number)
        
        # 4. Максимальное армирование записываем в ячейку A12 с большим шрифтом
        try:
            ws['A12'] = max_reinforcement
            ws['A12'].font = Font(size=48, bold=True)
            print(f"[FORMOVKA] ✅ Армирование {max_reinforcement} записано в ячейку A12 с размером шрифта 48")
        except Exception as e:
            print(f"[FORMOVKA] ⚠️ Не удалось записать армирование в A12: {e}")
        
        # 5. Заполняем плиты
        # Строка 11 - это заголовки таблицы (не трогаем)
        # Данные начинаются с 12 строки
        current_row = 12
        print(f"[FORMOVKA] 📝 Начинаю заполнять данные с 12 строки")
        for plate in plates_info:
            plate_name = plate.get('plate_name', '')
            qty = plate.get('qty', 0)
            kp_date = plate.get('kp_date', '')
            
            # Если нет готового имени плиты, формируем его
            if not plate_name or plate_name.strip() == '':
                length = plate.get('length', 0)
                width = plate.get('width', 1200)
                reinforcement = plate.get('reinforcement', 0)
                
                # Формируем имя плиты
                length_dm = int(round(length * 10))
                
                # Определяем нагрузку по армированию
                if reinforcement > 0:
                    if reinforcement < 8:
                        load_code = 6
                    elif reinforcement < 12:
                        load_code = 8
                    elif reinforcement < 15:
                        load_code = 10
                    else:
                        load_code = 12
                else:
                    load_code = 8
                
                # Форматируем ширину
                if width == 1200:
                    width_str = "12"
                else:
                    width_dm = width / 100.0
                    if abs(width_dm - int(width_dm)) < 0.01:
                        width_str = str(int(width_dm))
                    else:
                        width_str = str(width_dm).replace('.', ',')
                
                plate_name = f"ПБ {length_dm}-{width_str}-{load_code}п"
            
            # Формируем номер заказа (используем дату КП или текущую дату)
            if kp_date and kp_date != 'неизвестно':
                order_number = kp_date
            else:
                order_number = current_date
            
            # Колонка B - Заказ, №
            ws.cell(row=current_row, column=2).value = order_number
            
            # Колонка C - Номенклатура (название плиты)
            ws.cell(row=current_row, column=3).value = plate_name
            
            # Колонка D - Количество в формовку
            ws.cell(row=current_row, column=4).value = qty
            
            print(f"[FORMOVKA] Строка {current_row}: {plate_name} × {qty} шт, заказ {order_number}")
            current_row += 1
        
        # Сохраняем изменения
        wb.save(output_path)
        wb.close()
        
        print(f"[FORMOVKA] ✅ Файл формовки создан: {output_path}")
        return output_path
        
    except Exception as e:
        print(f"[FORMOVKA] ❌ Ошибка создания файла формовки: {e}")
        import traceback
        traceback.print_exc()
        return None


def create_formovka_files_for_tracks(
    tracks_data: List[Dict],
    output_dir: str,
    start_track_number: int = 1,
    template_path: str = "банк знаний/!КЗ ПБ Шаблон.xlsx",
    date_str: Optional[str] = None
) -> List[str]:
    """
    Создает Excel-файлы формовки для нескольких дорожек.
    
    Args:
        tracks_data: Список дорожек, каждая содержит:
            - track_number: номер дорожки (используется если есть)
            - max_reinforcement: максимальное армирование
            - plates_info: список плит на дорожке
        output_dir: Директория для сохранения файлов
        start_track_number: Начальный номер дорожки (используется если track_number не указан)
        template_path: Путь к шаблону
        date_str: Дата (по умолчанию текущая)
    
    Returns:
        Список путей к созданным файлам
    """
    created_files = []
    
    for idx, track_data in enumerate(tracks_data):
        # Используем номер дорожки из данных или вычисляем
        track_number = track_data.get('track_number', start_track_number + idx)
        max_reinforcement = track_data.get('max_reinforcement', 0)
        plates_info = track_data.get('plates_info', [])
        
        if not plates_info:
            print(f"[FORMOVKA] ⚠️ Пропускаю дорожку {track_number} - нет данных о плитах")
            continue
        
        file_path = create_formovka_excel(
            track_number=track_number,
            max_reinforcement=max_reinforcement,
            plates_info=plates_info,
            output_dir=output_dir,
            template_path=template_path,
            date_str=date_str
        )
        
        if file_path:
            created_files.append(file_path)
    
    print(f"[FORMOVKA] ✅ Создано {len(created_files)} файлов формовки")
    return created_files


# ==================== ТЕСТИРОВАНИЕ ====================

if __name__ == "__main__":
    # Тестовые данные
    test_plates = [
        {
            'plate_name': 'ПБ 80-12-10п',
            'qty': 4,
            'kp_date': '01.02.2026',
            'customer': 'Алексей'
        },
        {
            'plate_name': 'ПБ 56-12-6п',
            'qty': 2,
            'kp_date': '01.02.2026',
            'customer': 'Алексей'
        }
    ]
    
    # Создаем тестовый файл
    result = create_formovka_excel(
        track_number=1,
        max_reinforcement=39.0,
        plates_info=test_plates,
        output_dir="test_formovka",
        template_path="банк знаний/!КЗ ПБ Шаблон.xlsx"
    )
    
    if result:
        print(f"\n[OK] Тестовый файл формовки создан: {result}")
    else:
        print("\n[FAIL] Не удалось создать тестовый файл")
