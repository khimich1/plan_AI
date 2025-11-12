#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль генерации коммерческого предложения в формате PDF
Создаёт документ по образцу КП № 1133 от 16.10.2025
"""

import io
import os
import sqlite3
from datetime import datetime
from typing import List, Dict, Optional
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# ==================== КОНСТАНТЫ ====================

# Реквизиты компании
COMPANY_NAME = "ООО «Комбинат ЖБК»"
COMPANY_ADDRESS = "188300, Ленинградская область, г. Гатчина, ул. Заводская, д. 5"
COMPANY_PHONE = "+7 (812) 336-60-00"
COMPANY_EMAIL = "info@zhbk.ru"
COMPANY_INN = "4705123456"
COMPANY_KPP = "470501001"

# Банковские реквизиты
BANK_NAME = "ПАО Сбербанк"
BANK_BIK = "044030653"
BANK_ACCOUNT = "40702810123456789012"
BANK_CORR_ACCOUNT = "30101810500000000653"

# Путь к базе данных с ценами
DB_PATH = "pb.db"


# ==================== РЕГИСТРАЦИЯ ШРИФТОВ ====================

def register_fonts():
    """
    Регистрирует русские шрифты для ReportLab
    Ищет доступные шрифты Windows с поддержкой кириллицы
    """
    # Пути к стандартным шрифтам Windows
    windows_fonts = os.path.join(os.environ.get('WINDIR', 'C:\\Windows'), 'Fonts')
    
    # Список шрифтов для регистрации (имя в ReportLab, файл TTF)
    fonts_to_register = [
        ('DejaVuSans', 'DejaVuSans.ttf'),
        ('DejaVuSans-Bold', 'DejaVuSans-Bold.ttf'),
        ('Arial', 'arial.ttf'),
        ('Arial-Bold', 'arialbd.ttf'),
        ('TimesNewRoman', 'times.ttf'),
        ('TimesNewRoman-Bold', 'timesbd.ttf'),
    ]
    
    registered = False
    
    for font_name, font_file in fonts_to_register:
        font_path = os.path.join(windows_fonts, font_file)
        
        if os.path.exists(font_path):
            try:
                pdfmetrics.registerFont(TTFont(font_name, font_path))
                if not registered:
                    # Используем первый найденный шрифт как основной
                    globals()['FONT_NORMAL'] = font_name
                    globals()['FONT_BOLD'] = font_name + '-Bold' if font_name != 'DejaVuSans' else font_name
                    registered = True
            except Exception as e:
                continue
    
    if not registered:
        # Если не нашли ни одного TTF шрифта, используем встроенные
        # (но они не поддерживают кириллицу)
        globals()['FONT_NORMAL'] = 'Helvetica'
        globals()['FONT_BOLD'] = 'Helvetica-Bold'
    
    return registered


# Регистрируем шрифты при импорте модуля
HAS_CYRILLIC_FONTS = register_fonts()
FONT_NORMAL = globals().get('FONT_NORMAL', 'Helvetica')
FONT_BOLD = globals().get('FONT_BOLD', 'Helvetica-Bold')


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


def calculate_total_cost(order_data: List[Dict]) -> Dict:
    """
    Рассчитывает общую стоимость заказа
    
    Args:
        order_data: список позиций заказа с полями name, length_m, width_m, qty
    
    Returns:
        Словарь с итоговыми суммами
    """
    total_qty = 0
    total_cost = 0.0
    
    for item in order_data:
        qty = item.get('qty', 0)
        length_m = item.get('length_m', 0)
        width_m = item.get('width_m', 0)
        load_class = item.get('load_class', 800)
        
        # Получаем цену за единицу
        unit_price = get_plate_price(length_m, width_m, load_class)
        
        # Считаем сумму по позиции
        item_cost = unit_price * qty
        
        total_qty += qty
        total_cost += item_cost
    
    # НДС 20%
    vat_amount = round(total_cost * 0.20, 2)
    total_with_vat = round(total_cost + vat_amount, 2)
    
    return {
        'total_qty': total_qty,
        'subtotal': round(total_cost, 2),
        'vat_amount': vat_amount,
        'total_with_vat': total_with_vat
    }


def generate_commercial_offer_pdf(
    order_data: List[Dict],
    offer_number: str,
    offer_date: str,
    customer_name: Optional[str] = None
) -> io.BytesIO:
    """
    Генерирует коммерческое предложение в формате PDF
    
    Args:
        order_data: список позиций с полями:
            - name: название (например "ПБ 78-0.3-8п")
            - length_m: длина в метрах
            - width_m: ширина в метрах
            - qty: количество штук
            - load_class: класс нагрузки (опционально, по умолчанию 800)
        offer_number: номер коммерческого предложения
        offer_date: дата КП в формате "дд.мм.гггг"
        customer_name: название заказчика (опционально)
    
    Returns:
        BytesIO объект с содержимым PDF
    """
    
    # Создаём буфер для PDF
    buffer = io.BytesIO()
    
    # Создаём документ
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=20*mm,
        leftMargin=20*mm,
        topMargin=15*mm,
        bottomMargin=15*mm
    )
    
    # Стили
    styles = getSampleStyleSheet()
    
    # Кастомные стили с русскими шрифтами
    style_title = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontName=FONT_BOLD,
        fontSize=16,
        textColor=colors.HexColor('#1f4788'),
        spaceAfter=6*mm,
        alignment=1  # center
    )
    
    style_normal = ParagraphStyle(
        'CustomNormal',
        parent=styles['Normal'],
        fontName=FONT_NORMAL,
        fontSize=10,
        leading=14
    )
    
    style_small = ParagraphStyle(
        'CustomSmall',
        parent=styles['Normal'],
        fontName=FONT_NORMAL,
        fontSize=9,
        leading=12
    )
    
    # Элементы документа
    story = []
    
    # ==================== ШАПКА ====================
    
    # Логотип и название компании
    story.append(Paragraph(
        f"<b>{COMPANY_NAME}</b>",
        style_title
    ))
    
    story.append(Paragraph(
        f"{COMPANY_ADDRESS}<br/>"
        f"Тел.: {COMPANY_PHONE}, E-mail: {COMPANY_EMAIL}<br/>"
        f"ИНН {COMPANY_INN}, КПП {COMPANY_KPP}",
        style_small
    ))
    
    story.append(Spacer(1, 8*mm))
    
    # Заголовок документа
    story.append(Paragraph(
        f"<b>КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ № {offer_number}</b><br/>"
        f"от {offer_date}",
        style_title
    ))
    
    story.append(Spacer(1, 3*mm))
    
    # Заказчик (если указан)
    if customer_name:
        story.append(Paragraph(
            f"Заказчик: <b>{customer_name}</b>",
            style_normal
        ))
        story.append(Spacer(1, 3*mm))
    
    # Вводный текст
    story.append(Paragraph(
        "Уважаемые партнёры!",
        style_normal
    ))
    
    story.append(Spacer(1, 2*mm))
    
    story.append(Paragraph(
        f"{COMPANY_NAME} предлагает Вам железобетонные плиты перекрытий серии ПБ ЖБК СТАРТ "
        "собственного производства по следующим ценам:",
        style_normal
    ))
    
    story.append(Spacer(1, 5*mm))
    
    # ==================== ТАБЛИЦА С ПОЗИЦИЯМИ ====================
    
    # Заголовки таблицы
    table_data = [
        ['№', 'Наименование', 'Ед.изм.', 'Кол-во', 'Цена, руб.', 'Сумма, руб.']
    ]
    
    # Заполняем данные
    totals = calculate_total_cost(order_data)
    
    for idx, item in enumerate(order_data, start=1):
        name = item.get('name', 'Плита ПБ')
        qty = item.get('qty', 0)
        length_m = item.get('length_m', 0)
        width_m = item.get('width_m', 0)
        load_class = item.get('load_class', 800)
        
        # Получаем цену
        unit_price = get_plate_price(length_m, width_m, load_class)
        item_sum = unit_price * qty
        
        table_data.append([
            str(idx),
            name,
            'шт',
            str(qty),
            f"{unit_price:,.2f}".replace(',', ' '),
            f"{item_sum:,.2f}".replace(',', ' ')
        ])
    
    # Итоги
    table_data.append([
        '',
        'ИТОГО:',
        '',
        str(totals['total_qty']),
        '',
        f"{totals['subtotal']:,.2f}".replace(',', ' ')
    ])
    
    table_data.append([
        '',
        'НДС 20%:',
        '',
        '',
        '',
        f"{totals['vat_amount']:,.2f}".replace(',', ' ')
    ])
    
    table_data.append([
        '',
        'ВСЕГО с НДС:',
        '',
        '',
        '',
        f"{totals['total_with_vat']:,.2f}".replace(',', ' ')
    ])
    
    # Создаём таблицу
    table = Table(table_data, colWidths=[12*mm, 70*mm, 18*mm, 18*mm, 28*mm, 28*mm])
    
    # Стили таблицы
    table.setStyle(TableStyle([
        # Заголовок
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f4788')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), FONT_BOLD),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('TOPPADDING', (0, 0), (-1, 0), 8),
        
        # Данные
        ('BACKGROUND', (0, 1), (-1, -4), colors.white),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
        ('ALIGN', (0, 1), (0, -1), 'CENTER'),  # № по центру
        ('ALIGN', (3, 1), (3, -1), 'CENTER'),  # Кол-во по центру
        ('ALIGN', (4, 1), (-1, -1), 'RIGHT'),  # Цены справа
        ('FONTNAME', (0, 1), (-1, -1), FONT_NORMAL),
        ('FONTSIZE', (0, 1), (-1, -1), 9),
        ('GRID', (0, 0), (-1, -4), 0.5, colors.grey),
        
        # Итоги
        ('BACKGROUND', (0, -3), (-1, -1), colors.HexColor('#f0f0f0')),
        ('FONTNAME', (0, -3), (-1, -1), FONT_BOLD),
        ('FONTSIZE', (0, -3), (-1, -1), 10),
        ('LINEABOVE', (0, -3), (-1, -3), 1.5, colors.black),
        ('LINEABOVE', (0, -1), (-1, -1), 1.5, colors.black),
        ('LINEBELOW', (0, -1), (-1, -1), 1.5, colors.black),
        
        # Общие настройки
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ]))
    
    story.append(table)
    
    story.append(Spacer(1, 8*mm))
    
    # ==================== УСЛОВИЯ ====================
    
    story.append(Paragraph(
        "<b>Условия поставки:</b>",
        style_normal
    ))
    
    story.append(Spacer(1, 2*mm))
    
    conditions = [
        "• Срок изготовления: 5-7 рабочих дней с момента поступления оплаты",
        "• Форма оплаты: безналичный расчёт (100% предоплата)",
        "• Доставка: рассчитывается отдельно в зависимости от адреса и объёма",
        "• Разгрузка: силами и средствами заказчика",
        "• Срок действия предложения: 14 календарных дней"
    ]
    
    for condition in conditions:
        story.append(Paragraph(condition, style_small))
    
    story.append(Spacer(1, 5*mm))
    
    # ==================== БАНКОВСКИЕ РЕКВИЗИТЫ ====================
    
    story.append(Paragraph(
        "<b>Банковские реквизиты:</b>",
        style_normal
    ))
    
    story.append(Spacer(1, 2*mm))
    
    bank_details = [
        f"Получатель: {COMPANY_NAME}",
        f"ИНН {COMPANY_INN}, КПП {COMPANY_KPP}",
        f"Расчётный счёт: {BANK_ACCOUNT}",
        f"Банк: {BANK_NAME}",
        f"БИК: {BANK_BIK}",
        f"Корр. счёт: {BANK_CORR_ACCOUNT}"
    ]
    
    for detail in bank_details:
        story.append(Paragraph(detail, style_small))
    
    story.append(Spacer(1, 10*mm))
    
    # ==================== ПОДПИСЬ ====================
    
    story.append(Paragraph(
        "С уважением,<br/>"
        f"Отдел продаж {COMPANY_NAME}<br/>"
        f"Тел.: {COMPANY_PHONE}<br/>"
        f"E-mail: {COMPANY_EMAIL}",
        style_normal
    ))
    
    # Генерируем PDF
    doc.build(story)
    
    # Возвращаем буфер в начало
    buffer.seek(0)
    
    return buffer


# ==================== ТЕСТИРОВАНИЕ ====================

if __name__ == "__main__":
    # Тестовый заказ
    test_order = [
        {"name": "ПБ 56-6-8п", "length_m": 5.6, "width_m": 0.6, "qty": 1, "load_class": 800},
        {"name": "ПБ 78-0.3-8п", "length_m": 7.8, "width_m": 0.3, "qty": 3, "load_class": 800},
        {"name": "ПБ 68-6-8п", "length_m": 6.8, "width_m": 0.6, "qty": 2, "load_class": 800},
        {"name": "ПБ 56-9-8п", "length_m": 5.6, "width_m": 0.9, "qty": 1, "load_class": 800},
        {"name": "ПБ 78-9-8п", "length_m": 7.8, "width_m": 0.9, "qty": 3, "load_class": 800},
        {"name": "ПБ 52-11-8п", "length_m": 5.2, "width_m": 1.1, "qty": 1, "load_class": 800},
    ]
    
    # Генерируем PDF
    pdf_buffer = generate_commercial_offer_pdf(
        order_data=test_order,
        offer_number="1133",
        offer_date=datetime.now().strftime("%d.%m.%Y"),
        customer_name="ООО «Тестовая компания»"
    )
    
    # Сохраняем в файл
    output_path = "test_commercial_offer.pdf"
    with open(output_path, 'wb') as f:
        f.write(pdf_buffer.getvalue())
    
    print(f"[OK] Test KP created: {output_path}")

