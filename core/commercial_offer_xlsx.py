#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль генерации коммерческого предложения в формате XLSX
Создаёт документ по образцу КП с расчётными формулами
"""

import io
import logging
import os
from typing import List, Dict, Optional

try:
    import pandas as pd
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.drawing.image import Image as XLImage
    HAS_EXCEL_LIBS = True
except ImportError:
    HAS_EXCEL_LIBS = False
    XLImage = None
    logging.getLogger(__name__).warning(
        "pandas или openpyxl не установлены; генерация XLSX будет недоступна"
    )

# Единый расчёт веса строки КП (plate_weights → approximate)
try:
    from .cargo_delivery_pricing import (
        cargo_delivery_trips_count,
        delivery_service_charge_rub,
        total_order_cargo_weight_kg,
    )
    from .kp_plate_weight import resolve_kp_line_weight_kg
except ImportError:
    from cargo_delivery_pricing import (
        cargo_delivery_trips_count,
        delivery_service_charge_rub,
        total_order_cargo_weight_kg,
    )
    from kp_plate_weight import resolve_kp_line_weight_kg


# ==================== КОНСТАНТЫ ====================

# Вычисляем абсолютный путь к корню проекта (папка выше core/)
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Реквизиты компании
COMPANY_NAME = "ООО «Комбинат ЖБК»"
COMPANY_ADDRESS = "150020, г Ярославль г, проезд Домостроителей, дом 1, строение 3"
COMPANY_PHONE = "8 (4852) 595 000"
COMPANY_EMAIL = "info@zhbk.ru"

# Путь к базе данных с ценами (абсолютный)
DB_PATH = os.path.join(PROJECT_ROOT, "pb.db")

logger = logging.getLogger(__name__)

try:
    from core.project_paths import resolve_commercial_offer_logo_path
except ImportError:
    from project_paths import resolve_commercial_offer_logo_path

# Коэффициент скидки в процентах (0 = без скидки, 5 = скидка 5%, и т.д.)
DISCOUNT_PERCENT = 0


# ==================== ФУНКЦИИ ====================

def get_plate_price(length_m: float, width_m: float, load_class: int = 800) -> float:
    """
    Получает цену плиты из базы данных по длине, ширине и классу нагрузки
    
    Args:
        length_m: длина плиты в метрах
        width_m: ширина плиты в метрах  
        load_class: класс нагрузки (по умолчанию 800 кг/м²)
    
    Returns:
        Цена плиты в рублях
    
    Raises:
        PriceNotFoundError: если цена отсутствует в прайс-листе
    """
    from core.commercial_pricing import lookup_plate_price

    return lookup_plate_price(length_m, width_m, load_class, db_path=DB_PATH)


from core.commercial_pricing import (  # noqa: E402
    VAT_RATE,
    calculate_total_cost as _calculate_total_cost,
    format_phone,
    is_pile_order,
    lookup_pile_price,
)


def calculate_total_cost(
    order_data: List[Dict], discount_percent: float = 0, logistics_cost: float = 0
) -> Dict:
    return _calculate_total_cost(
        order_data,
        discount_percent,
        logistics_cost,
        db_path=DB_PATH,
    )


def generate_commercial_offer_xlsx(
    order_data: List[Dict],
    offer_number: str,
    offer_date: str,
    customer_name: Optional[str] = None,
    manager_name: Optional[str] = None,
    manager_phone: Optional[str] = None,
    manager_email: Optional[str] = None,
    discount_percent: float = 0,
    delivery_conditions: Optional[str] = None,
    payment_conditions: Optional[str] = None,
    kp_db_id: Optional[int] = None,
    logistics_cost: float = 0.0,
) -> io.BytesIO:
    """
    Генерирует коммерческое предложение в формате XLSX с расчётными формулами
    
    Args:
        order_data: список позиций заказа с полями name, length_m, width_m, qty
        offer_number: номер КП
        offer_date: дата КП
        customer_name: имя заказчика (будет в строке 16)
        manager_name: имя менеджера (выводится в подписи)
        manager_phone: телефон менеджера (выводится под именем)
        manager_email: email менеджера (выводится под телефоном)
        discount_percent: процент скидки (0-100, по умолчанию 0)
        delivery_conditions: условия поставки (строка 28, если указано)
        payment_conditions: условия оплаты (строка 29, если указано)
        kp_db_id: номер КП из базы данных (если КП сохранен в БД)
        logistics_cost: стоимость одного рейса (без НДС по плитам; итог по формуле рейсов)
    
    Returns:
        BytesIO буфер с XLSX файлом
    """
    if not HAS_EXCEL_LIBS:
        raise ImportError("Для генерации XLSX необходимы pandas и openpyxl")
    
    buffer = io.BytesIO()
    
    # Определяем, есть ли логотип (плашка; будет вставлен выше заголовка)
    logo_path = resolve_commercial_offer_logo_path()
    has_logo = logo_path is not None
    
    # Формируем заголовок документа
    # Если есть логотип, добавляем пустые строки для него
    header_data = []
    
    if has_logo:
        # Резервируем место под логотип (7 строк как в образце)
        header_data.extend([[""], [""], [""], [""], [""], [""], [""]])
    
    # Формируем строку с номером КП и датой
    if kp_db_id:
        date_line = f"КП № {kp_db_id} от {offer_date}"
    else:
        date_line = f"от {offer_date}"
    
    # Формируем строку с контактами (МЕНЕДЖЕРА, а не компании!)
    if manager_phone and manager_email:
        contacts_line = f"Тел.: {format_phone(manager_phone)}, email: {manager_email}"
    elif manager_phone:
        contacts_line = f"Тел.: {format_phone(manager_phone)}"
    elif manager_email:
        contacts_line = f"Email: {manager_email}"
    else:
        # Fallback: контакты компании, если нет контактов менеджера
        contacts_line = f"Тел.: {COMPANY_PHONE}, Email: {COMPANY_EMAIL}"
    
    header_data.extend([
        [f"{COMPANY_ADDRESS}"],
        [contacts_line],  # 🆕 КОНТАКТЫ МЕНЕДЖЕРА (не компании!)
        [""],
        [date_line],
        ["Срок действия коммерческого предложения 3 дня"],
        [""],
        ["Коммерческое предложение для"],
        [customer_name or "Заказчик не указан"],
        [""],
    ])
    
    # Создаем DataFrame для заголовка
    df_header = pd.DataFrame(header_data)
    
    # Формируем таблицу товаров
    table_data = []
    pile_order = is_pile_order(order_data)
    if pile_order:
        table_headers = ['№', 'Наименование', 'Класс бетона', 'Кол-во', 'Цена', 'Сумма']
    else:
        table_headers = ['№', 'Наименование', 'Кол-во', 'Ед.', 'Вес(кг)', 'Цена', 'Сумма']
    
    total_weight = 0.0
    
    for idx, item in enumerate(order_data, start=1):
        qty = item.get('qty', 0)

        if 'unit_price' in item and item['unit_price'] is not None:
            unit_price = item['unit_price']
        elif pile_order:
            mark = str(item.get('mark') or item.get('name') or '').strip()
            grade = str(item.get('concrete_grade') or 'B25').strip()
            unit_price = lookup_pile_price(mark, grade, db_path=DB_PATH)
        else:
            name = item.get('name', 'Плиты ПБ')
            length_m = item.get('length_m', 0)
            width_m = item.get('width_m', 0)
            load_class = item.get('load_class')
            if load_class is None:
                try:
                    from config_and_data import parse_load_code_from_name
                    load_code = parse_load_code_from_name(name, default=8)
                    load_class = max(1, load_code) * 100
                except ImportError:
                    load_class = 800
            unit_price = get_plate_price(length_m, width_m, load_class)

        discounted_price = unit_price * (1 - discount_percent / 100)

        if pile_order:
            table_data.append({
                '№': idx,
                'Наименование': str(item.get('mark') or item.get('name') or ''),
                'Класс бетона': str(item.get('concrete_grade') or 'B25'),
                'Кол-во': qty,
                'Цена': discounted_price,
                'Сумма': discounted_price * qty,
            })
            continue

        name = item.get('name', 'Плиты ПБ')
        _, total_item_weight = resolve_kp_line_weight_kg(item)
        total_weight += total_item_weight
        table_data.append({
            '№': idx,
            'Наименование': name,
            'Кол-во': qty,
            'Ед.': 'шт',
            'Вес(кг)': total_item_weight,
            'Цена': discounted_price,
            'Сумма': discounted_price * qty
        })
    
    trip_cost = max(0.0, float(logistics_cost or 0.0))
    delivery_trips = cargo_delivery_trips_count(total_weight)
    has_delivery_line = (not pile_order) and trip_cost > 0 and delivery_trips > 0
    if has_delivery_line:
        delivery_total = delivery_service_charge_rub(trip_cost, total_weight)
        table_data.append(
            {
                '№': len(table_data) + 1,
                'Наименование': 'Услуга по доставке грузов',
                'Кол-во': delivery_trips,
                'Ед.': 'рейс',
                'Вес(кг)': 0.0,
                'Цена': trip_cost,
                'Сумма': delivery_total,
            }
        )

    df_table = pd.DataFrame(table_data)
    
    # Рассчитываем итоги с учётом скидки
    totals = calculate_total_cost(order_data, discount_percent, logistics_cost=logistics_cost)
    
    # Создаем строку итогов
    total_items = len(table_data)
    
    # Записываем в Excel
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        # Записываем заголовок (без заголовков столбцов)
        df_header.to_excel(writer, sheet_name='КП', index=False, header=False, startrow=0)
        
        # Записываем таблицу товаров (с заголовками)
        start_row = len(header_data)
        df_table.to_excel(writer, sheet_name='КП', index=False, startrow=start_row)
        
        # Получаем workbook и worksheet для форматирования
        workbook = writer.book
        worksheet = writer.sheets['КП']
        
        # === ВСТАВКА ЛОГОТИПА ===
        if has_logo and XLImage is not None:
            try:
                # Создаем объект изображения
                logo_img = XLImage(str(logo_path))
                
                # Масштабируем логотип (оригинал 1456x207)
                # Делаем ширину больше для лучшей читаемости
                logo_img.width = 1000  # пикселей (было 600)
                logo_img.height = int(logo_img.width * (207 / 1456.0))  # сохраняем пропорции
                
                # Объединяем ячейки для логотипа (7 строк x 7 колонок, как в образце)
                worksheet.merge_cells('A1:G7')
                
                # Вставляем в ячейку A1
                worksheet.add_image(logo_img, 'A1')
                
                # Увеличиваем высоту первых 7 строк для логотипа
                total_logo_height = logo_img.height
                for row_idx in range(1, 8):
                    worksheet.row_dimensions[row_idx].height = total_logo_height / 7
                    
            except Exception as e:
                logger.warning("Не удалось вставить логотип в XLSX: %s", e)
        
        # === ФОРМАТИРОВАНИЕ ===
        
        # Шрифты и стили
        header_font = Font(name='Tahoma', size=11, bold=True)
        title_font = Font(name='Tahoma', size=18, bold=True)
        customer_font = Font(name='Tahoma', size=16, bold=True)
        table_header_font = Font(name='Tahoma', size=11, bold=True)
        table_font = Font(name='Tahoma', size=10)
        summary_font = Font(name='Tahoma', size=11, bold=True)
        
        center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
        left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
        right_align = Alignment(horizontal='right', vertical='center')
        
        # Границы для таблицы
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Заливка для заголовка таблицы
        header_fill = PatternFill(start_color='F0F0F0', end_color='F0F0F0', fill_type='solid')
        
        # Определяем смещение строк (если есть логотип, добавляем 7 строк)
        logo_offset = 7 if has_logo else 0
        
        # Форматируем заголовок документа (с учетом смещения для логотипа)
        company_row = logo_offset + 1
        address_row = logo_offset + 2
        phone_row = logo_offset + 3
        number_row = logo_offset + 5
        validity_row = logo_offset + 6
        title_row = logo_offset + 8
        customer_row = logo_offset + 9
        
        worksheet[f'A{company_row}'].font = header_font
        worksheet[f'A{address_row}'].font = header_font
        worksheet[f'A{phone_row}'].font = header_font
        worksheet[f'A{number_row}'].font = header_font
        worksheet[f'A{validity_row}'].font = header_font
        worksheet[f'A{title_row}'].font = title_font
        worksheet[f'A{title_row}'].alignment = center_align
        worksheet[f'A{customer_row}'].font = customer_font
        worksheet[f'A{customer_row}'].alignment = center_align
        
        # Объединяем ячейки для заголовка
        worksheet.merge_cells(f'A{company_row}:G{company_row}')
        worksheet.merge_cells(f'A{address_row}:G{address_row}')
        worksheet.merge_cells(f'A{phone_row}:G{phone_row}')
        worksheet.merge_cells(f'A{number_row}:G{number_row}')
        worksheet.merge_cells(f'A{validity_row}:G{validity_row}')
        worksheet.merge_cells(f'A{title_row}:G{title_row}')
        worksheet.merge_cells(f'A{customer_row}:G{customer_row}')
        
        # Форматируем заголовок таблицы
        table_header_row = start_row + 1
        for col_idx in range(1, 8):  # 7 столбцов
            cell = worksheet.cell(row=table_header_row, column=col_idx)
            cell.font = table_header_font
            cell.alignment = center_align
            cell.fill = header_fill
            cell.border = thin_border
        
        # Форматируем строки таблицы
        table_rows_count = len(table_data)
        for row_idx in range(table_header_row + 1, table_header_row + 1 + table_rows_count):
            # №
            worksheet.cell(row=row_idx, column=1).alignment = center_align
            worksheet.cell(row=row_idx, column=1).border = thin_border
            worksheet.cell(row=row_idx, column=1).font = table_font
            
            # Наименование
            worksheet.cell(row=row_idx, column=2).alignment = left_align
            worksheet.cell(row=row_idx, column=2).border = thin_border
            worksheet.cell(row=row_idx, column=2).font = table_font
            
            # Кол-во
            worksheet.cell(row=row_idx, column=3).alignment = center_align
            worksheet.cell(row=row_idx, column=3).border = thin_border
            worksheet.cell(row=row_idx, column=3).font = table_font
            
            # Ед.
            worksheet.cell(row=row_idx, column=4).alignment = center_align
            worksheet.cell(row=row_idx, column=4).border = thin_border
            worksheet.cell(row=row_idx, column=4).font = table_font
            
            # Вес
            worksheet.cell(row=row_idx, column=5).alignment = right_align
            worksheet.cell(row=row_idx, column=5).border = thin_border
            worksheet.cell(row=row_idx, column=5).font = table_font
            worksheet.cell(row=row_idx, column=5).number_format = '#,##0.00'
            
            # Цена
            worksheet.cell(row=row_idx, column=6).alignment = right_align
            worksheet.cell(row=row_idx, column=6).border = thin_border
            worksheet.cell(row=row_idx, column=6).font = table_font
            worksheet.cell(row=row_idx, column=6).number_format = '#,##0.00'
            
            # Сумма - добавляем ФОРМУЛУ!
            sum_cell = worksheet.cell(row=row_idx, column=7)
            sum_cell.value = f"=C{row_idx}*F{row_idx}"  # Кол-во * Цена
            sum_cell.alignment = right_align
            sum_cell.border = thin_border
            sum_cell.font = table_font
            sum_cell.number_format = '#,##0.00'
        
        # Добавляем итоговые строки
        summary_row = table_header_row + table_rows_count + 2
        
        # Итого по таблице: сумма позиций (цены с НДС) + услуга доставки при наличии.
        subtotal_row = summary_row
        worksheet.merge_cells(f'A{subtotal_row}:F{subtotal_row}')
        worksheet[f'A{subtotal_row}'] = f"Всего наименований {total_items}, общим весом {total_weight:,.3f} кг."
        worksheet[f'A{subtotal_row}'].font = summary_font
        worksheet[f'A{subtotal_row}'].alignment = left_align

        first_data_row = table_header_row + 1
        last_data_row = table_header_row + table_rows_count
        subtotal_cell = worksheet[f'G{subtotal_row}']
        subtotal_cell.value = f"=SUM(G{first_data_row}:G{last_data_row})"
        subtotal_cell.font = summary_font
        subtotal_cell.number_format = '#,##0.00'
        subtotal_cell.alignment = right_align

        # НДС 22% только от суммы плит (без строки «Услуга по доставке грузов»).
        last_plate_row = last_data_row - 1 if has_delivery_line else last_data_row
        vat_row = summary_row + 1
        worksheet.merge_cells(f'A{vat_row}:F{vat_row}')
        worksheet[f'A{vat_row}'] = "в том числе НДС (22%)"
        worksheet[f'A{vat_row}'].font = summary_font
        worksheet[f'A{vat_row}'].alignment = left_align
        worksheet[f'G{vat_row}'] = f"=SUM(G{first_data_row}:G{last_plate_row})*{VAT_RATE}"
        worksheet[f'G{vat_row}'].font = summary_font
        worksheet[f'G{vat_row}'].number_format = '#,##0.00'
        worksheet[f'G{vat_row}'].alignment = right_align
        
        # Условия
        conditions_row = vat_row + 2
        
        # Условия поставки (строка 28 с учётом смещения)
        if delivery_conditions:
            worksheet[f'A{conditions_row}'] = f"1. Условия поставки: {delivery_conditions}"
        else:
            worksheet[f'A{conditions_row}'] = "1. Условия поставки:"
        worksheet[f'A{conditions_row}'].font = table_font
        
        conditions_row += 1
        
        # Условия оплаты (строка 29 с учётом смещения)
        if payment_conditions:
            worksheet[f'A{conditions_row}'] = f"2. Условия оплаты: {payment_conditions}"
        else:
            worksheet[f'A{conditions_row}'] = "2. Условия оплаты: Предварительная оплата в размере 100%"
        worksheet[f'A{conditions_row}'].font = table_font
        
        # Подпись
        signature_row = conditions_row + 2
        worksheet[f'A{signature_row}'] = "С уважением,"
        worksheet[f'A{signature_row}'].font = summary_font
        
        signature_row += 1
        # Используем только имя менеджера (телефон и email теперь в шапке)
        worksheet[f'A{signature_row}'] = manager_name or "Менеджер"
        worksheet[f'A{signature_row}'].font = table_font
        
        # Примечание
        note_row = signature_row + 2
        worksheet[f'A{note_row}'] = ('Доборные плиты ПБ отгружаются только при наличии в наименовании "+доб", '
                                      'для получения доборов, просим сообщить Вашему менеджеру до начала изготовления. '
                                      'Все доборы по умолчанию отправляем на утилизацию.')
        worksheet[f'A{note_row}'].font = Font(name='Tahoma', size=9)
        worksheet[f'A{note_row}'].alignment = left_align
        worksheet.merge_cells(f'A{note_row}:G{note_row}')
        
        # Настройка ширины столбцов
        worksheet.column_dimensions['A'].width = 5
        worksheet.column_dimensions['B'].width = 45
        if pile_order:
            worksheet.column_dimensions['C'].width = 14
            worksheet.column_dimensions['D'].width = 8
            worksheet.column_dimensions['E'].width = 15
            worksheet.column_dimensions['F'].width = 18
        else:
            worksheet.column_dimensions['C'].width = 10
            worksheet.column_dimensions['D'].width = 8
            worksheet.column_dimensions['E'].width = 15
            worksheet.column_dimensions['F'].width = 18
            worksheet.column_dimensions['G'].width = 18
        
        # Высота строк
        worksheet.row_dimensions[table_header_row].height = 25
    
    buffer.seek(0)
    return buffer

