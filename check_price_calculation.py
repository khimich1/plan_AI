#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Проверка расчета итоговой цены плиты
"""
import sys
import os

# Устанавливаем кодировку для вывода
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Импортируем нужные модули
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from price_db import get_price
from config_and_data import PRICE_DB_PATH, LONG_CUT_PRICE_PER_M, TRANSVERSE_CUT_PRICE

print("=" * 70)
print("ПРОВЕРКА РАСЧЕТА ИТОГОВОЙ ЦЕНЫ ПЛИТЫ")
print("=" * 70)

# Параметры плиты из изображения
length_m = 3.6
width_m = 0.32  # 320 мм
load_code = 8

print(f"\nПараметры плиты:")
print(f"  Длина: {length_m} м")
print(f"  Ширина: {width_m} м (320 мм)")
print(f"  Класс нагрузки: {load_code}")

# 1. Базовая цена из базы данных
base_price_1_2m = get_price(length_m, load_code, PRICE_DB_PATH)
print(f"\n1. БАЗОВАЯ ЦЕНА (из базы данных):")
print(f"   Цена плиты 1.2м × {length_m}м: {base_price_1_2m:,.2f} руб")

# 2. Сколько плит получается из одной исходной плиты
# Для плиты 320 мм из плиты 1200 мм: 1200 / 320 = 3.75, но обычно получается 3 плиты + остаток
# Но в данном случае из изображения видно, что из одной плиты получается 1 плита 320мм + остаток 880мм
plates_from_source = 1  # По изображению - 1 плита из одной исходной
base_price = base_price_1_2m / plates_from_source
print(f"   Плит из одной исходной: {plates_from_source}")
print(f"   → Базовая цена плиты: {base_price_1_2m:,.2f} / {plates_from_source} = {base_price:,.2f} руб")

# 3. Стоимость резов
# По изображению: 1 продольный рез
long_cuts = 1
trans_cuts = 0
cuts_cost = long_cuts * (LONG_CUT_PRICE_PER_M * length_m) + trans_cuts * TRANSVERSE_CUT_PRICE
print(f"\n2. СТОИМОСТЬ РЕЗОВ:")
print(f"   Продольных резов: {long_cuts}")
print(f"   Поперечных резов: {trans_cuts}")
print(f"   → Стоимость резов: {long_cuts} × (460 руб/м × {length_m}м) + {trans_cuts} × 1200 руб = {cuts_cost:,.2f} руб")

# 4. Стоимость остатков (нужно проверить в коде)
# По изображению остаток 880 мм не использован
rest_cost = 0  # Пока неизвестно, нужно проверить расчет
print(f"\n3. СТОИМОСТЬ НЕИСПОЛЬЗОВАННЫХ ОСТАТКОВ:")
print(f"   (требуется проверка расчета)")

# 5. Стоимость отходов
waste_cost = 0  # По изображению отходов нет
print(f"\n4. СТОИМОСТЬ ОТХОДОВ:")
print(f"   {waste_cost:,.2f} руб")

# Итоговая цена
unit_price = base_price + cuts_cost + rest_cost + waste_cost
print(f"\n" + "=" * 70)
print(f"ИТОГОВАЯ ЦЕНА (без остатков и отходов):")
print(f"   {base_price:,.2f} + {cuts_cost:,.2f} = {unit_price:,.2f} руб")
print(f"\nЦЕНА НА ИЗОБРАЖЕНИИ: 21,993.20 руб")
print(f"РАЗНИЦА: {21993.20 - unit_price:,.2f} руб")
print("=" * 70)

print(f"\nВЫВОД:")
print(f"Если итоговая цена 21,993.20 руб, а базовая + резы = {unit_price:,.2f} руб,")
print(f"то стоимость остатков и отходов должна быть: {21993.20 - unit_price:,.2f} руб")
print(f"\nЭто означает, что остаток 880 мм учитывается в стоимости!")


