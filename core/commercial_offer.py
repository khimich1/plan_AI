#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль генерации коммерческого предложения в формате PDF
Создаёт документ по образцу КП № 1133 от 16.10.2025

Для генерации XLSX используйте модуль commercial_offer_xlsx
"""

import io
import os
import sqlite3
from datetime import datetime
from typing import List, Dict, Optional
from xml.sax.saxutils import escape
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

# Вычисляем абсолютный путь к корню проекта (папка выше core/)
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

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

# Путь к базе данных с ценами (абсолютный)
DB_PATH = os.path.join(PROJECT_ROOT, "pb.db")

# Путь к логотипу (абсолютный, чтобы работал из любой директории)
LOGO_PATH = os.path.join(PROJECT_ROOT, "банк знаний", "ЖБЛСТАРТ.png")

# Коэффициент скидки в процентах (0 = без скидки, 5 = скидка 5%, и т.д.)
DISCOUNT_PERCENT = 0


# ==================== РЕГИСТРАЦИЯ ШРИФТОВ ====================

def register_fonts():
    """
    Регистрирует русские шрифты для ReportLab и старается выбрать Tahoma,
    чтобы повторить фирменное оформление КП.
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    font_paths = []
    local_font_dirs = [
        os.path.join(script_dir, 'fonts'),
        os.path.join(script_dir, 'банк знаний'),
        os.path.join(script_dir, 'bank'),
    ]
    
    for local_dir in local_font_dirs:
        if os.path.exists(local_dir):
            font_paths.append(local_dir)
    
    linux_font_paths = [
        '/usr/share/fonts/truetype/dejavu/',
        '/usr/share/fonts/truetype/liberation/',
        '/usr/share/fonts/truetype/msttcorefonts/',
        '/usr/share/fonts/truetype/freefont/',
        '/usr/share/fonts/TTF/',
        '/usr/share/fonts/',
        '~/.fonts/',
    ]
    
    windows_fonts = os.path.join(os.environ.get('WINDIR', 'C:\\Windows'), 'Fonts')
    if os.path.exists(windows_fonts):
        font_paths.append(windows_fonts)
    
    for path in linux_font_paths:
        expanded = os.path.expanduser(path)
        if os.path.exists(expanded):
            font_paths.append(expanded)
    
    # Убираем дубликаты, сохраняя порядок
    seen_paths = set()
    unique_font_paths = []
    for path in font_paths:
        if path not in seen_paths:
            unique_font_paths.append(path)
            seen_paths.add(path)
    
    fonts_to_register = [
        ('Tahoma', ('Tahoma.ttf', 'tahoma.ttf')),
        ('Tahoma-Bold', ('tahomabd.ttf', 'Tahoma-Bold.ttf', 'Tahoma Bold.ttf', 'tahoma-bold.ttf')),
        ('DejaVuSans', ('DejaVuSans.ttf',)),
        ('DejaVuSans-Bold', ('DejaVuSans-Bold.ttf',)),
        ('LiberationSans', ('LiberationSans-Regular.ttf', 'LiberationSans.ttf')),
        ('LiberationSans-Bold', ('LiberationSans-Bold.ttf',)),
        ('Arial', ('Arial.ttf', 'arial.ttf')),
        ('Arial-Bold', ('Arial-Bold.ttf', 'arialbd.ttf')),
        ('TimesNewRoman', ('TimesNewRoman.ttf', 'times.ttf')),
        ('TimesNewRoman-Bold', ('TimesNewRoman-Bold.ttf', 'timesbd.ttf')),
    ]
    
    registered_fonts = set()
    
    for font_name, candidates in fonts_to_register:
        for font_dir in unique_font_paths:
            for candidate in candidates:
                font_path = os.path.join(font_dir, candidate)
                if os.path.exists(font_path):
                    try:
                        pdfmetrics.registerFont(TTFont(font_name, font_path))
                        registered_fonts.add(font_name)
                        break
                    except Exception:
                        continue
            if font_name in registered_fonts:
                break
    
    preferred_pairs = [
        ('Tahoma', 'Tahoma-Bold'),
        ('DejaVuSans', 'DejaVuSans-Bold'),
        ('LiberationSans', 'LiberationSans-Bold'),
        ('Arial', 'Arial-Bold'),
        ('TimesNewRoman', 'TimesNewRoman-Bold'),
    ]
    
    for normal, bold in preferred_pairs:
        if normal in registered_fonts:
            globals()['FONT_NORMAL'] = normal
            globals()['FONT_BOLD'] = bold if bold in registered_fonts else normal
            break
    else:
        if registered_fonts:
            fallback = next(iter(registered_fonts))
            globals()['FONT_NORMAL'] = fallback
            globals()['FONT_BOLD'] = fallback
        else:
            globals()['FONT_NORMAL'] = 'Helvetica'
            globals()['FONT_BOLD'] = 'Helvetica-Bold'
            print("ВНИМАНИЕ: Не найдены шрифты с поддержкой кириллицы! Русский текст может отображаться некорректно.")
    
    return bool(registered_fonts)


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
        order_data: список позиций заказа с полями name, length_m, width_m, qty, unit_price (опционально)
    
    Returns:
        Словарь с итоговыми суммами
    """
    total_qty = 0
    total_cost = 0.0
    
    for item in order_data:
        qty = item.get('qty', 0)

        # 🔥 ПРИОРИТЕТ: Если цена уже рассчитана (с учётом резов/отходов), используем её!
        if 'unit_price' in item and item['unit_price'] is not None:
            unit_price = item['unit_price']
        else:
            # Fallback: старая логика (только базовая цена из БД)
            length_m = item.get('length_m', 0)
            width_m = item.get('width_m', 0)
            load_class = item.get('load_class', 800)
            unit_price = get_plate_price(length_m, width_m, load_class)

        # Применяем скидку к цене (если указана)
        discounted_price = unit_price * (1 - DISCOUNT_PERCENT / 100)

        # Считаем сумму по позиции
        item_cost = discounted_price * qty
        
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
    Генерирует коммерческое предложение в формате PDF, повторяя фирменное
    оформление КП № 1133 от 16.10.2025.
    
    Args:
        order_data: список позиций заказа
        offer_number: номер коммерческого предложения
        offer_date: дата формирования
        customer_name: имя заказчика
    
    Note:
        Детальная разбивка компонентов НЕ включается в PDF.
        Она сохраняется в отдельный Excel файл через save_breakdown_to_excel()
    """
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=10 * mm,
        leftMargin=10 * mm,
        topMargin=12 * mm,
        bottomMargin=15 * mm
    )
    
    content_width = doc.width
    styles = getSampleStyleSheet()
    
    style_header_info = ParagraphStyle(
        'HeaderInfo',
        parent=styles['Normal'],
        fontName=FONT_BOLD,
        fontSize=11,
        leading=14,
        spaceAfter=2 * mm
    )
    
    style_doc_number = ParagraphStyle(
        'DocNumber',
        parent=styles['Normal'],
        fontName=FONT_BOLD,
        fontSize=11,
        leading=14,
        spaceAfter=4 * mm
    )
    
    style_title = ParagraphStyle(
        'OfferTitle',
        parent=styles['Normal'],
        fontName=FONT_BOLD,
        fontSize=18,
        leading=22,
        alignment=1,
        spaceAfter=1 * mm
    )
    
    style_customer = ParagraphStyle(
        'OfferCustomer',
        parent=styles['Normal'],
        fontName=FONT_BOLD,
        fontSize=16,
        leading=20,
        alignment=1,
        spaceAfter=6 * mm
    )
    
    style_table_text = ParagraphStyle(
        'TableText',
        parent=styles['Normal'],
        fontName=FONT_NORMAL,
        fontSize=10,
        leading=12
    )
    
    style_summary = ParagraphStyle(
        'Summary',
        parent=styles['Normal'],
        fontName=FONT_BOLD,
        fontSize=11,
        leading=15,
        spaceAfter=6 * mm
    )
    
    style_conditions = ParagraphStyle(
        'Conditions',
        parent=styles['Normal'],
        fontName=FONT_NORMAL,
        fontSize=11,
        leading=14,
        spaceAfter=2 * mm
    )
    
    style_signature_label = ParagraphStyle(
        'SignatureLabel',
        parent=styles['Normal'],
        fontName=FONT_BOLD,
        fontSize=12,
        leading=16
    )
    
    style_signature = ParagraphStyle(
        'Signature',
        parent=styles['Normal'],
        fontName=FONT_NORMAL,
        fontSize=12,
        leading=16
    )
    
    style_note = ParagraphStyle(
        'Note',
        parent=styles['Normal'],
        fontName=FONT_NORMAL,
        fontSize=9,
        leading=12,
        textColor=colors.HexColor('#333333'),
        spaceBefore=2 * mm
    )
    
    story = []
    
    if os.path.exists(LOGO_PATH):
        try:
            logo_width = content_width
            logo_height = logo_width * (207 / 1456.0)
            story.append(Image(LOGO_PATH, width=logo_width, height=logo_height))
            story.append(Spacer(1, 4 * mm))
        except Exception as exc:
            print(f"Ошибка загрузки логотипа: {exc}")
    
    story.append(Paragraph(
        f"{COMPANY_ADDRESS}<br/>Тел.: {COMPANY_PHONE}",
        style_header_info
    ))
    
    story.append(Paragraph(
        f"№ {offer_number} от {offer_date}",
        style_doc_number
    ))
    
    story.append(Paragraph(
        "Коммерческое предложение для",
        style_title
    ))
    
    customer_display = escape(customer_name.strip()) if customer_name else "Заказчик не указан"
    story.append(Paragraph(customer_display, style_customer))
    
    story.append(Spacer(1, 6 * mm))
    
    table_data = [['№', 'Наименование', 'Кол-во', 'Ед.', 'Вес(кг)', 'Цена', 'Сумма']]
    
    totals = calculate_total_cost(order_data)
    total_weight = 0.0
    
    for idx, item in enumerate(order_data, start=1):
        name_raw = item.get('name', 'Плиты ПБ')
        name = escape(name_raw)
        qty = item.get('qty', 0)
        length_m = item.get('length_m', 0)
        width_m = item.get('width_m', 0)
        load_class = item.get('load_class')

        # 🔥 ПРИОРИТЕТ: Если цена уже рассчитана (с учётом резов/отходов), используем её!
        if 'unit_price' in item and item['unit_price'] is not None:
            unit_price = item['unit_price']
        else:
            # Fallback: старая логика (только базовая цена из БД)
            # Если класс нагрузки явно не передан, пробуем вытащить его из имени плиты.
            # Формат имени: "Плиты ПБ 71-12-10п", "ПБ 69-12-12,5п" и т.п.
            if load_class is None:
                try:
                    from config_and_data import parse_load_code_from_name
                except ImportError:
                    load_class = 800
                else:
                    load_code = parse_load_code_from_name(name_raw, default=8)
                    load_class = max(1, load_code) * 100  # 8 -> 800, 10 -> 1000 и т.п.
            
            unit_price = get_plate_price(length_m, width_m, load_class)
        
        # Применяем скидку к цене (если указана)
        discounted_price = unit_price * (1 - DISCOUNT_PERCENT / 100)
        item_sum = discounted_price * qty
        
        # Вес: если уже передан в item, используем его, иначе рассчитываем
        if 'weight' in item and item['weight'] is not None:
            total_item_weight = item['weight']
        else:
            unit_weight = approximate_weight_kg(length_m, width_m)
            total_item_weight = unit_weight * qty
        
        total_weight += total_item_weight
        
        weight_str = f"{total_item_weight:,.2f}".replace(',', 'X').replace('.', ',').replace('X', ' ')
        price_str = f"{discounted_price:,.2f}".replace(',', 'X').replace('.', ',').replace('X', ' ')
        sum_str = f"{item_sum:,.2f}".replace(',', 'X').replace('.', ',').replace('X', ' ')
        
        table_data.append([
            str(idx),
            Paragraph(name, style_table_text),
            str(qty),
            'шт',
            weight_str,
            price_str,
            sum_str
        ])
    
    no_width = 10 * mm
    qty_width = 14 * mm
    unit_width = 11 * mm
    weight_width = 20 * mm
    price_width = 25 * mm
    sum_width = 25 * mm
    
    fixed_total = no_width + qty_width + unit_width + weight_width + price_width + sum_width
    name_width = content_width - fixed_total
    
    if name_width <= 0:
        ratios = (0.05, 0.45, 0.08, 0.06, 0.18, 0.09, 0.09)
        col_widths = [content_width * ratio for ratio in ratios]
    else:
        col_widths = [
            no_width,
            name_width,
            qty_width,
            unit_width,
            weight_width,
            price_width,
            sum_width
        ]
    
    table = Table(table_data, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#f0f0f0')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('FONTNAME', (0, 0), (-1, 0), FONT_BOLD),
        ('FONTSIZE', (0, 0), (-1, 0), 11),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('TOPPADDING', (0, 0), (-1, 0), 6),
        
        ('FONTNAME', (0, 1), (-1, -1), FONT_NORMAL),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('ALIGN', (0, 1), (0, -1), 'CENTER'),
        ('ALIGN', (2, 1), (3, -1), 'CENTER'),
        ('ALIGN', (4, 1), (-1, -1), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 5),
        ('TOPPADDING', (0, 1), (-1, -1), 5),
        
        ('BOX', (0, 0), (-1, -1), 1.2, colors.black),
        ('INNERGRID', (0, 0), (-1, -1), 0.6, colors.black),
    ]))
    
    story.append(table)
    story.append(Spacer(1, 6 * mm))
    
    # ❌ ДЕТАЛЬНАЯ РАЗБИВКА НЕ ДОБАВЛЯЕТСЯ В PDF
    # Она сохраняется в отдельный Excel файл и отправляется вместе с PDF
    
    total_items = len(order_data)
    weight_summary = f"{total_weight:,.3f}".replace(',', 'X').replace('.', ',').replace('X', ' ')
    subtotal_str = f"{totals['subtotal']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', ' ')
    vat_str = f"{totals['vat_amount']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', ' ')
    
    summary_text = (
        f"Всего наименований {total_items}, общим весом {weight_summary} кг., "
        f"на сумму {subtotal_str} руб., в том числе НДС {vat_str} руб."
    )
    story.append(Paragraph(summary_text, style_summary))
    
    story.append(Paragraph("1.Условия поставки:", style_conditions))
    story.append(Paragraph("2.Условия оплаты: Предварительная оплата в размере 100%", style_conditions))
    
    story.append(Spacer(1, 6 * mm))
    
    story.append(Paragraph("С уважением,", style_signature_label))
    story.append(Paragraph("Шишов Александр Васильевич", style_signature))
    story.append(Paragraph("8 920 640 55 85", style_signature))
    
    story.append(Paragraph(
        'Доборные плиты ПБ отгружаются только при наличии в наименовании "+доб", '
        "для получения доборов, просим сообщить Вашему менеджеру до начала изготовления.<br/>"
        "Все доборы по умолчанию отправляем на утилизацию.",
        style_note
    ))
    
    doc.build(story)
    buffer.seek(0)
    return buffer


def save_breakdown_to_excel(breakdown_tables: List[Dict], output_path: str) -> bool:
    """
    Сохраняет детальную разбивку компонентов в отдельный Excel файл.
    
    Args:
        breakdown_tables: список таблиц с детальной разбивкой
        output_path: путь для сохранения Excel файла
    
    Returns:
        True если успешно сохранено, False если ошибка
    """
    try:
        import pandas as pd
    except ImportError:
        print("[BREAKDOWN] pandas не установлен, невозможно создать Excel файл")
        return False
    
    if not breakdown_tables:
        print("[BREAKDOWN] breakdown_tables пустой, нечего сохранять")
        return False
    
    try:
        breakdown_headers = ['Компонент', 'Расчёт', 'Сумма']
        all_breakdown_rows = []
        
        for breakdown in breakdown_tables:
            # Заголовок с наименованием плиты
            all_breakdown_rows.append([breakdown['name'], '', ''])
            
            # Строки таблицы (все 3 столбца)
            for row in breakdown['rows']:
                # Берём все 3 столбца
                all_breakdown_rows.append(row if len(row) >= 3 else row + [''] * (3 - len(row)))
            
            # Пустая строка между таблицами
            all_breakdown_rows.append(['', '', ''])
        
        # Удаляем последнюю пустую строку
        if all_breakdown_rows and all_breakdown_rows[-1] == ['', '', '']:
            all_breakdown_rows.pop()
        
        df_breakdown = pd.DataFrame(all_breakdown_rows, columns=breakdown_headers)
        df_breakdown.to_excel(output_path, index=False)
        print(f'[BREAKDOWN] ✅ Детальная разбивка сохранена: {output_path}')
        return True
        
    except Exception as e:
        print(f'[BREAKDOWN] ❌ Ошибка сохранения: {e}')
        import traceback
        traceback.print_exc()
        return False


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

