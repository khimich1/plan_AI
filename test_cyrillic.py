#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Тест поддержки кириллицы в PDF"""

from commercial_offer import generate_commercial_offer_pdf, HAS_CYRILLIC_FONTS, FONT_NORMAL, FONT_BOLD
from datetime import datetime

print("="*60)
print("Test: Cyrillic Support in PDF")
print("="*60)

print(f"\n[INFO] Cyrillic fonts available: {HAS_CYRILLIC_FONTS}")
print(f"[INFO] Normal font: {FONT_NORMAL}")
print(f"[INFO] Bold font: {FONT_BOLD}")

# Тестовый заказ с русскими названиями
test_order = [
    {"name": "Плита ПБ 78-3-8п", "length_m": 7.8, "width_m": 0.3, "qty": 3, "load_class": 800},
    {"name": "Плита ПБ 56-6-8п", "length_m": 5.6, "width_m": 0.6, "qty": 1, "load_class": 800},
    {"name": "Плита ПБ 68-6-8п", "length_m": 6.8, "width_m": 0.6, "qty": 2, "load_class": 800},
]

try:
    pdf_buffer = generate_commercial_offer_pdf(
        order_data=test_order,
        offer_number="TEST_1133",
        offer_date=datetime.now().strftime("%d.%m.%Y"),
        customer_name="ООО «Тестовая компания»"
    )
    
    output_path = "test_cyrillic_kp.pdf"
    with open(output_path, 'wb') as f:
        f.write(pdf_buffer.getvalue())
    
    print(f"\n[OK] PDF with cyrillic generated successfully!")
    print(f"[OK] File: {output_path}")
    print(f"[OK] Size: {len(pdf_buffer.getvalue())} bytes")
    print(f"\n[SUCCESS] Please check the file for correct Russian text display")
    
except Exception as e:
    print(f"\n[ERROR] Failed: {e}")
    import traceback
    traceback.print_exc()

print("="*60)

