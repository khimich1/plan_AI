#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модули визуализации:
- price_utils: работа с ценами
- procurement: построение закупки и сметы
- layout_sequence: построение последовательности раскладки
- visualization_drawing: функции рисования
"""

# Реэкспорт основных функций для удобства импорта
from .price_utils import (
    load_price_table_from_xlsx,
    sync_price_xlsx_to_db,
    find_price_from_db,
    find_price_for_plate,
    load_cut_price_from_docx
)

from .procurement import (
    get_orders_from_opt_plan,
    build_procurement_items,
    build_price_rows,
    build_component_breakdown,
    build_price_rows_production,
    build_component_breakdown_production
)

from .layout_sequence import build_layout_sequence

from .visualization_drawing import (
    _draw_segment,
    _draw_split_plate,
    _draw_transverse_cut
)

__all__ = [
    # price_utils
    'load_price_table_from_xlsx',
    'sync_price_xlsx_to_db',
    'find_price_from_db',
    'find_price_for_plate',
    'load_cut_price_from_docx',
    # procurement
    'get_orders_from_opt_plan',
    'build_procurement_items',
    'build_price_rows',
    'build_component_breakdown',
    'build_price_rows_production',
    'build_component_breakdown_production',
    # layout_sequence
    'build_layout_sequence',
    # visualization_drawing
    '_draw_segment',
    '_draw_split_plate',
    '_draw_transverse_cut',
]

