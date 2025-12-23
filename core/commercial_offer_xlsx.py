#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль генерации коммерческого предложения в формате XLSX
Создаёт документ по образцу КП с расчётными формулами
"""

import io
import os
import sqlite3
from typing import List, Dict, Optional

try:
    import pandas as pd
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.drawing.image import Image as XLImage
    HAS_EXCEL_LIBS = True
except ImportError:
    HAS_EXCEL_LIBS = False
    XLImage = None
    print("ВНИМАНИЕ: pandas или openpyxl не установлены. Генерация XLSX будет недоступна.")

# Импортируем функцию расчёта веса
try:
    from config_and_data import approximate_weight_kg
except ImportError:
    # Если не удалось импортировать, используем локальную функцию
    def approximate_weight_kg(length_m: float, width_m: float, thickness_m: float = 0.22) -> float:
        """Примерный расчёт веса плиты в килограммах"""
        volume = length_m * width_m * thickness_m
        return round(volume * 2400, 2)


# ==================== КОНСТАНТЫ ====================

# Реквизиты компании
COMPANY_NAME = "ООО «Комбинат ЖБК»"
COMPANY_ADDRESS = "150020, г Ярославль г, проезд Домостроителей, дом 1, строение 3"
COMPANY_PHONE = "8 (4852) 595 000"
COMPANY_EMAIL = "info@zhbk.ru"

# Путь к базе данных с ценами
DB_PATH = "pb.db"

# Путь к логотипу
LOGO_PATH = "банк знаний/ЖБЛСТАРТ.png"

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
    """
    try:
        # Преобразуем в дециметры для поиска в базе
        length_dm = int(round(length_m * 10))
        
        # Определяем код нагрузки (8 = 800 кг/м², 10 = 1000 кг/м²)
        load_code = load_class // 100
        
        con = sqlite3.connect(DB_PATH)
        cur = con.cursor()
        
        # Ищем цену в таблице prices
        result = cur.execute(
            "SELECT price FROM prices WHERE length_dm = ? AND load_code = ?",
            (length_dm, load_code)
        ).fetchone()
        
        con.close()
        
        if result:
            return float(result[0])
        else:
            # Если нет точной цены, используем базовую формулу
            # Примерная цена: 4000 руб/м² * площадь плиты
            area_m2 = length_m * width_m
            return round(area_m2 * 4000, 2)
            
    except Exception as e:
        print(f"Ошибка получения цены: {e}")
        # Возвращаем примерную цену
        area_m2 = length_m * width_m
        return round(area_m2 * 4000, 2)


def calculate_total_cost(order_data: List[Dict], discount_percent: float = 0) -> Dict:
    """
    Рассчитывает общую стоимость заказа
    
    Args:
        order_data: список позиций заказа с полями name, length_m, width_m, qty, unit_price (опционально)
        discount_percent: процент скидки (0-100)
    
    Returns:
        Словарь с итоговыми суммами
    """
    total_qty = 0
    total_cost_with_vat = 0.0  # Сумма с НДС (unit_price уже включает НДС)
    
    for item in order_data:
        qty = item.get('qty', 0)
        
        # 🔥 ПРИОРИТЕТ: Если цена уже рассчитана (с учётом резов/отходов), используем её!
        if 'unit_price' in item and item['unit_price'] is not None:
            unit_price = item['unit_price']  # Цена УЖЕ включает НДС
        else:
            # Fallback: старая логика (только базовая цена из БД)
            length_m = item.get('length_m', 0)
            width_m = item.get('width_m', 0)
            load_class = item.get('load_class', 800)
            unit_price = get_plate_price(length_m, width_m, load_class)
        
        # Применяем скидку к цене (если указана)
        discounted_price = unit_price * (1 - discount_percent / 100)
        
        # Считаем сумму по позиции (это уже с НДС)
        item_cost = discounted_price * qty
        
        total_qty += qty
        total_cost_with_vat += item_cost
    
    # 🔥 ИСПРАВЛЕНИЕ: unit_price уже включает НДС, поэтому нужно вычесть НДС
    # Сумма без НДС = сумма с НДС / 1.20
    subtotal = round(total_cost_with_vat / 1.20, 2)
    vat_amount = round(total_cost_with_vat - subtotal, 2)
    total_with_vat = round(total_cost_with_vat, 2)
    
    return {
        'total_qty': total_qty,
        'subtotal': subtotal,
        'vat_amount': vat_amount,
        'total_with_vat': total_with_vat
    }


def generate_commercial_offer_xlsx(
    order_data: List[Dict],
    offer_number: str,
    offer_date: str,
    customer_name: Optional[str] = None,
    manager_name: Optional[str] = None,
    discount_percent: float = 0,
    delivery_conditions: Optional[str] = None,
    payment_conditions: Optional[str] = None
) -> io.BytesIO:
    """
    Генерирует коммерческое предложение в формате XLSX с расчётными формулами
    
    Args:
        order_data: список позиций заказа с полями name, length_m, width_m, qty
        offer_number: номер КП
        offer_date: дата КП
        customer_name: имя заказчика (будет в строке 16)
        manager_name: имя менеджера (будет в ячейке A31)
        discount_percent: процент скидки (0-100, по умолчанию 0)
        delivery_conditions: условия поставки (строка 28, если указано)
        payment_conditions: условия оплаты (строка 29, если указано)
    
    Returns:
        BytesIO буфер с XLSX файлом
    """
    if not HAS_EXCEL_LIBS:
        raise ImportError("Для генерации XLSX необходимы pandas и openpyxl")
    
    buffer = io.BytesIO()
    
    # Определяем, есть ли логотип (будет вставлен выше заголовка)
    has_logo = os.path.exists(LOGO_PATH)
    
    # Формируем заголовок документа
    # Если есть логотип, добавляем пустые строки для него
    header_data = []
    
    if has_logo:
        # Резервируем место под логотип (7 строк как в образце)
        header_data.extend([[""], [""], [""], [""], [""], [""], [""]])
    
    header_data.extend([
        [COMPANY_NAME],
        [f"{COMPANY_ADDRESS}"],
        [f"Тел.: {COMPANY_PHONE}, Email: {COMPANY_EMAIL}"],
        [""],
        [f"от {offer_date}"],
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
    table_headers = ['№', 'Наименование', 'Кол-во', 'Ед.', 'Вес(кг)', 'Цена', 'Сумма']
    
    total_weight = 0.0
    
    for idx, item in enumerate(order_data, start=1):
        name = item.get('name', 'Плиты ПБ')
        qty = item.get('qty', 0)
        length_m = item.get('length_m', 0)
        width_m = item.get('width_m', 0)
        load_class = item.get('load_class')
        
        # 🔥 ПРИОРИТЕТ: Если цена уже рассчитана (с учётом резов/отходов), используем её!
        if 'unit_price' in item and item['unit_price'] is not None:
            unit_price = item['unit_price']
        else:
            # Fallback: старая логика (только базовая цена из БД)
            # Определяем класс нагрузки из имени, если не передан
            if load_class is None:
                try:
                    from config_and_data import parse_load_code_from_name
                    load_code = parse_load_code_from_name(name, default=8)
                    load_class = max(1, load_code) * 100
                except ImportError:
                    load_class = 800
            
            unit_price = get_plate_price(length_m, width_m, load_class)
        
        # Вес: если уже передан в item, используем его, иначе рассчитываем
        if 'weight' in item and item['weight'] is not None:
            unit_weight = item['weight'] / qty  # Общий вес делим на количество
        else:
            unit_weight = approximate_weight_kg(length_m, width_m)
        
        total_item_weight = unit_weight * qty
        total_weight += total_item_weight
        
        # Применяем скидку к цене (если указана)
        discounted_price = unit_price * (1 - discount_percent / 100)
        
        table_data.append({
            '№': idx,
            'Наименование': name,
            'Кол-во': qty,
            'Ед.': 'шт',
            'Вес(кг)': total_item_weight,
            'Цена': discounted_price,
            'Сумма': discounted_price * qty  # Сначала вставим значение, потом заменим на формулу
        })
    
    df_table = pd.DataFrame(table_data)
    
    # Рассчитываем итоги с учётом скидки
    totals = calculate_total_cost(order_data, discount_percent)
    
    # Создаем строку итогов
    total_items = len(order_data)
    
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
                logo_img = XLImage(LOGO_PATH)
                
                # Масштабируем логотип (оригинал 1456x207)
                # Делаем ширину примерно на всю ширину таблицы (7 колонок)
                logo_img.width = 600  # пикселей
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
                print(f"[XLSX] Не удалось вставить логотип: {e}")
        
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
        for row_idx in range(table_header_row + 1, table_header_row + 1 + len(order_data)):
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
        summary_row = table_header_row + len(order_data) + 2
        
        # Если есть скидка, добавляем строку с информацией о скидке
        if discount_percent > 0:
            discount_row = summary_row
            summary_row = discount_row + 1  # Сдвигаем итоги на одну строку вниз
            
            worksheet.merge_cells(f'A{discount_row}:F{discount_row}')
            discount_cell = worksheet[f'A{discount_row}']
            discount_cell.value = f"Скидка {discount_percent}%"
            discount_cell.font = Font(name='Tahoma', size=11, bold=True, color='FF006100')  # Зелёный цвет для скидки
            discount_cell.alignment = left_align
        
        # Итоговая сумма (уже с НДС, так как unit_price включает НДС)
        subtotal_row = summary_row
        worksheet.merge_cells(f'A{subtotal_row}:F{subtotal_row}')
        worksheet[f'A{subtotal_row}'] = f"Всего наименований {total_items}, общим весом {total_weight:,.3f} кг."
        worksheet[f'A{subtotal_row}'].font = summary_font
        worksheet[f'A{subtotal_row}'].alignment = left_align
        
        # Формула для подсчёта суммы (это сумма с НДС)
        first_data_row = table_header_row + 1
        last_data_row = table_header_row + len(order_data)
        sum_with_vat_cell = worksheet[f'G{subtotal_row}']
        sum_with_vat_cell.value = f"=SUM(G{first_data_row}:G{last_data_row})"
        sum_with_vat_cell.font = summary_font
        sum_with_vat_cell.number_format = '#,##0.00'
        sum_with_vat_cell.alignment = right_align
        
        # НДС 20% - вычитаем из суммы с НДС
        vat_row = summary_row + 1
        worksheet.merge_cells(f'A{vat_row}:F{vat_row}')
        worksheet[f'A{vat_row}'] = "в том числе НДС (20%)"
        worksheet[f'A{vat_row}'].font = summary_font
        worksheet[f'A{vat_row}'].alignment = left_align
        # НДС = сумма с НДС - сумма без НДС = сумма с НДС - (сумма с НДС / 1.20)
        worksheet[f'G{vat_row}'] = f"=G{subtotal_row}-G{subtotal_row}/1.2"  # Формула для НДС
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
        # Используем имя менеджера из параметров или значение по умолчанию
        worksheet[f'A{signature_row}'] = manager_name or "Шишов Александр Васильевич"
        worksheet[f'A{signature_row}'].font = table_font
        
        signature_row += 1
        worksheet[f'A{signature_row}'] = "8 920 640 55 85"
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
        worksheet.column_dimensions['C'].width = 10
        worksheet.column_dimensions['D'].width = 8
        worksheet.column_dimensions['E'].width = 15
        worksheet.column_dimensions['F'].width = 18
        worksheet.column_dimensions['G'].width = 18
        
        # Высота строк
        worksheet.row_dimensions[table_header_row].height = 25
    
    buffer.seek(0)
    return buffer

