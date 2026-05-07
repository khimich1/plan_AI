# -*- coding: utf-8 -*-
"""Публичный срез API парсинга заказов плит."""

from .plate_lists import (
    LineContributionKey,
    get_last_parse_diagnostics,
    set_plate_lists_from_text,
)

__all__ = [
    "LineContributionKey",
    "get_last_parse_diagnostics",
    "set_plate_lists_from_text",
]
