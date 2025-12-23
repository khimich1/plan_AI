#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль расчёта заводской себестоимости плит ПБ

Основные компоненты:
- db_schema.py: Схема БД для хранения себестоимости
- excel_reader.py: Чтение данных из Excel
- import_from_xlsx.py: Импорт себестоимости в БД
- cost_engine.py: API для получения себестоимости

ВАЖНО: Модуль НЕ работает с КП и таблицей prices.
Это изолированная система для расчёта ЗАВОДСКОЙ себестоимости.
"""

from .cost_engine import (
    get_cost_by_plate_name,
    get_cost_by_params,
)

__all__ = [
    'get_cost_by_plate_name',
    'get_cost_by_params',
]

