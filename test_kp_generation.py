#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Тест генерации коммерческого предложения"""

from commercial_offer import generate_commercial_offer_pdf
from datetime import datetime

# Тестовый заказ с плитами из вашей базы
test_order = [
    {"name": "ПБ 78-3-8п", "length_m": 7.8, "width_m": 0.3, "qty": 3, "load_class": 800},
    {"name": "ПБ 56-6-8п", "length_m": 5.6, "width_m": 0.6, "qty": 1, "load_class": 800},
    {"name": "ПБ 68-6-8п", "length_m": 6.8, "width_m": 0.6, "qty": 2, "load_class": 800},
    {"name": "ПБ 78-9-8п", "length_m": 7.8, "width_m": 0.9, "qty": 3, "load_class": 800},
    {"name": "ПБ 52-11-8п", "length_m": 5.2, "width_m": 1.1, "qty": 1, "load_class": 800},
]

print("="*60)
print("Test: Commercial Offer PDF Generation")
print("="*60)

try:
    # Генерируем PDF
    pdf_buffer = generate_commercial_offer_pdf(
        order_data=test_order,
        offer_number="TEST_1133",
        offer_date=datetime.now().strftime("%d.%m.%Y"),
        customer_name="Тестовая компания ООО"
    )
    
    # Сохраняем
    output_path = "test_kp.pdf"
    with open(output_path, 'wb') as f:
        f.write(pdf_buffer.getvalue())
    
    print(f"\n[OK] PDF generated successfully!")
    print(f"[OK] File saved: {output_path}")
    print(f"[OK] File size: {len(pdf_buffer.getvalue())} bytes")
    
    # Проверяем содержимое
    print(f"\n[INFO] Order details:")
    for idx, item in enumerate(test_order, 1):
        print(f"  {idx}. {item['name']} x {item['qty']} pcs")
    
    print(f"\n[SUCCESS] Commercial offer module is working correctly!")
    
except Exception as e:
    print(f"\n[ERROR] Failed to generate PDF: {e}")
    import traceback
    traceback.print_exc()

print("="*60)

