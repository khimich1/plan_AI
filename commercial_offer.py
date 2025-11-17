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
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

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

# Реквизиты компании (согласно примеру КП)
COMPANY_NAME = "ООО «Комбинат ЖБК»"
COMPANY_ADDRESS = "150020, г Ярославль г, проезд Домостроителей, дом 1, строение 3"
COMPANY_PHONE = "8 (4852) 595 000"
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

# Путь к логотипу
LOGO_PATH = "банк знаний/ЖБЛСТАРТ.png"


# ==================== РЕГИСТРАЦИЯ ШРИФТОВ ====================

def register_fonts():
    """
    Регистрирует русские шрифты для ReportLab
    Ищет доступные шрифты в системе (Linux/Windows) с поддержкой кириллицы
    """
    # Список путей для поиска шрифтов
    font_paths = []
    
    # Linux пути
    linux_font_paths = [
        '/usr/share/fonts/truetype/dejavu/',
        '/usr/share/fonts/truetype/liberation/',
        '/usr/share/fonts/TTF/',
        '/usr/share/fonts/',
        '~/.fonts/',
    ]
    
    # Windows пути
    windows_fonts = os.path.join(os.environ.get('WINDIR', 'C:\\Windows'), 'Fonts')
    if os.path.exists(windows_fonts):
        font_paths.append(windows_fonts)
    
    # Добавляем Linux пути
    for path in linux_font_paths:
        expanded = os.path.expanduser(path)
        if os.path.exists(expanded):
            font_paths.append(expanded)
    
    # Список шрифтов для регистрации (имя в ReportLab, файл TTF)
    fonts_to_register = [
        ('DejaVuSans', 'DejaVuSans.ttf'),
        ('DejaVuSans-Bold', 'DejaVuSans-Bold.ttf'),
        ('LiberationSans', 'LiberationSans-Regular.ttf'),
        ('LiberationSans-Bold', 'LiberationSans-Bold.ttf'),
        ('Arial', 'arial.ttf'),
        ('Arial-Bold', 'arialbd.ttf'),
        ('TimesNewRoman', 'times.ttf'),
        ('TimesNewRoman-Bold', 'timesbd.ttf'),
    ]
    
    registered_normal = False
    registered_bold = False
    
    # Сначала регистрируем все шрифты
    for font_name, font_file in fonts_to_register:
        for font_dir in font_paths:
            font_path = os.path.join(font_dir, font_file)
            
            if os.path.exists(font_path):
                try:
                    pdfmetrics.registerFont(TTFont(font_name, font_path))
                    if font_name.endswith('-Bold'):
                        registered_bold = True
                    else:
                        registered_normal = True
                except Exception as e:
                    continue
    
    # Выбираем основной шрифт
    if registered_normal:
        # Ищем обычный шрифт
        for font_name, font_file in fonts_to_register:
            if font_name.endswith('-Bold'):
                continue
            for font_dir in font_paths:
                font_path = os.path.join(font_dir, font_file)
                if os.path.exists(font_path):
                    globals()['FONT_NORMAL'] = font_name
                    # Ищем соответствующий жирный шрифт
                    if font_name == 'DejaVuSans' and registered_bold:
                        globals()['FONT_BOLD'] = 'DejaVuSans-Bold'
                    elif font_name == 'LiberationSans' and registered_bold:
                        globals()['FONT_BOLD'] = 'LiberationSans-Bold'
                    elif font_name.startswith('Arial') and registered_bold:
                        globals()['FONT_BOLD'] = 'Arial-Bold'
                    elif font_name.startswith('TimesNewRoman') and registered_bold:
                        globals()['FONT_BOLD'] = 'TimesNewRoman-Bold'
                    else:
                        globals()['FONT_BOLD'] = font_name  # Используем обычный, если жирный не найден
                    break
            if 'FONT_NORMAL' in globals():
                break
    
    registered = registered_normal
    
    if not registered:
        # Если не нашли ни одного TTF шрифта, используем встроенные
        # (но они не поддерживают кириллицу)
        globals()['FONT_NORMAL'] = 'Helvetica'
        globals()['FONT_BOLD'] = 'Helvetica-Bold'
        print("ВНИМАНИЕ: Не найдены шрифты с поддержкой кириллицы! Русский текст может отображаться некорректно.")
    
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
    
    # Создаём документ (отступы точь в точь как в образце)
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=10*mm,  # Еще больше уменьшено - контент ближе к правому краю
        leftMargin=10*mm,   # Еще больше уменьшено - контент ближе к левому краю
        topMargin=5*mm,     # Логотип ближе к верху
        bottomMargin=15*mm
    )
    
    # Ширина контента (A4 ширина минус отступы)
    content_width = A4[0] - 20*mm  # 210mm - 20mm = 190mm
    
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
    
    # Логотип на всю ширину листа (если файл существует)
    logo_path = LOGO_PATH
    if os.path.exists(logo_path):
        try:
            # Логотип на всю ширину контента
            # Исходный размер: 1456x207 пикселей
            logo_width = content_width
            logo_height = (207 / 1456) * logo_width  # Пропорциональная высота
            logo = Image(logo_path, width=logo_width, height=logo_height)
            story.append(logo)
            story.append(Spacer(1, 2*mm))  # Уменьшено для соответствия образцу
        except Exception as e:
            print(f"Ошибка загрузки логотипа: {e}")
    
    # Реквизиты компании (как в образце: адрес и телефон)
    story.append(Paragraph(
        f"{COMPANY_ADDRESS}<br/>"
        f"Тел.: {COMPANY_PHONE}",
        style_small
    ))
    
    story.append(Spacer(1, 2*mm))  # Уменьшено для соответствия образцу
    
    # Заголовок документа (как в образце: "№ 1133 от 16.10.2025")
    story.append(Paragraph(
        f"№ {offer_number} от {offer_date}",
        style_normal
    ))
    
    story.append(Spacer(1, 1.5*mm))  # Уменьшено для соответствия образцу
    
    # Заголовок "Коммерческое предложение для" (отдельная строка как в образце)
    story.append(Paragraph(
        "Коммерческое предложение для",
        style_normal
    ))
    
    # Имя заказчика (отдельная строка как в образце)
    if customer_name:
        story.append(Paragraph(
            f"<b>{customer_name}</b>",
            style_normal
        ))
    else:
        story.append(Paragraph(
            "<b>[Заказчик не указан]</b>",
            style_normal
        ))
    
    story.append(Spacer(1, 4*mm))  # Уменьшено для соответствия образцу
    
    # ==================== ТАБЛИЦА С ПОЗИЦИЯМИ ====================
    
    # Заголовки таблицы (как в примере: № Наименование Кол-во Ед. Вес(кг) Цена Сумма)
    table_data = [
        ['№', 'Наименование', 'Кол-во', 'Ед.', 'Вес(кг)', 'Цена', 'Сумма']
    ]
    
    # Заполняем данные
    totals = calculate_total_cost(order_data)
    total_weight = 0.0
    
    for idx, item in enumerate(order_data, start=1):
        name = item.get('name', 'Плита ПБ')
        qty = item.get('qty', 0)
        length_m = item.get('length_m', 0)
        width_m = item.get('width_m', 0)
        load_class = item.get('load_class', 800)
        
        # Получаем цену
        unit_price = get_plate_price(length_m, width_m, load_class)
        item_sum = unit_price * qty
        
        # Рассчитываем вес одной плиты
        unit_weight = approximate_weight_kg(length_m, width_m)
        total_item_weight = unit_weight * qty
        total_weight += total_item_weight
        
        # Форматируем числа: пробелы для тысяч, запятая для десятичных
        weight_str = f"{total_item_weight:,.2f}".replace(',', 'X').replace('.', ',').replace('X', ' ')
        price_str = f"{unit_price:,.2f}".replace(',', 'X').replace('.', ',').replace('X', ' ')
        sum_str = f"{item_sum:,.2f}".replace(',', 'X').replace('.', ',').replace('X', ' ')
        table_data.append([
            str(idx),
            name,
            str(qty),
            'шт',
            weight_str,
            price_str,
            sum_str
        ])
    
    # Итоги (не добавляем в таблицу, выведем отдельно как в примере)
    
    # Создаём таблицу (ширины колонок: №, Наименование, Кол-во, Ед., Вес(кг), Цена, Сумма)
    table = Table(table_data, colWidths=[10*mm, 60*mm, 15*mm, 12*mm, 20*mm, 25*mm, 25*mm])
    
    # Стили таблицы (как в образце - светлый серый фон заголовка)
    table.setStyle(TableStyle([
        # Заголовок (светлый серый фон, черный текст - как в образце)
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#e0e0e0')),  # Светлый серый
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),  # Черный текст
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), FONT_BOLD),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('TOPPADDING', (0, 0), (-1, 0), 8),
        
        # Данные
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
        ('ALIGN', (0, 1), (0, -1), 'CENTER'),  # № по центру
        ('ALIGN', (2, 1), (2, -1), 'CENTER'),  # Кол-во по центру
        ('ALIGN', (3, 1), (3, -1), 'CENTER'),  # Ед. по центру
        ('ALIGN', (4, 1), (4, -1), 'RIGHT'),   # Вес справа
        ('ALIGN', (5, 1), (-1, -1), 'RIGHT'),  # Цены справа
        ('FONTNAME', (0, 1), (-1, -1), FONT_NORMAL),
        ('FONTSIZE', (0, 1), (-1, -1), 9),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        
        # Общие настройки
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ]))
    
    story.append(table)
    
    story.append(Spacer(1, 5*mm))
    
    # Итоги (как в примере: "Всего наименований 10, общим весом 19 826,875 кг., на сумму 185 764,80 руб., в том числе НДС 30 960,79 руб.")
    total_items = len(order_data)
    # Форматируем числа с запятой для десятичных и пробелами для тысяч
    weight_str = f"{total_weight:,.3f}".replace(',', 'X').replace('.', ',').replace('X', ' ')
    subtotal_str = f"{totals['subtotal']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', ' ')
    vat_str = f"{totals['vat_amount']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', ' ')
    story.append(Paragraph(
        f"Всего наименований {total_items}, общим весом {weight_str} кг., "
        f"на сумму {subtotal_str} руб., в том числе НДС {vat_str} руб.",
        style_normal
    ))
    
    story.append(Spacer(1, 8*mm))
    
    # ==================== УСЛОВИЯ ====================
    
    # Формат как в образце: только заголовки без деталей
    story.append(Paragraph(
        "<b>1.Условия поставки:</b>",
        style_normal
    ))
    
    story.append(Spacer(1, 3*mm))
    
    story.append(Paragraph(
        "<b>2.Условия оплаты:</b> Предварительная оплата в размере 100%",
        style_normal
    ))
    
    story.append(Spacer(1, 5*mm))
    
    # ==================== ПОДПИСЬ ====================
    
    # Формат как в примере: "С уважением, Шишов Александр Васильевич 8 920 640 55 85"
    story.append(Paragraph(
        "С уважением,<br/>"
        "Шишов Александр Васильевич<br/>"
        "8 920 640 55 85",
        style_normal
    ))
    
    story.append(Spacer(1, 5*mm))
    
    # Примечание о доборных плитах (как в примере)
    story.append(Paragraph(
        "Доборные плиты ПБ отгружаются только при наличии в наименовании \"+доб\", "
        "для получения доборов, просим сообщить Вашему менеджеру до начала изготовления.<br/>"
        "Все доборы по умолчанию отправляем на утилизацию.",
        style_small
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

