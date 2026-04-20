#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Быстрый интеграционный тест"""

from factory_cost import get_cost_by_params
from core.config import set_plate_lists_from_text

# Парсим заказ
set_plate_lists_from_text('ПБ 71-12-10п 5 шт')

# Получаем себестоимость
cost = get_cost_by_params(7.1, 1.2)

if cost:
    print(f"OK: {cost['plate_name']} = {cost['full_cost_with_kef']:.2f} руб")
else:
    print("FAIL: Плита не найдена")

