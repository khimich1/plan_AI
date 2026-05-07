#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Настройки из окружения (чтение при импорте; не менять в рантайме)."""

from __future__ import annotations

import os

# Источник веса для КП:
# - "formula": расчет по формуле (дм * дм * коэффициент)
# - "plate_weights": legacy-режим через таблицу plate_weights
WEIGHT_SOURCE = os.getenv("WEIGHT_SOURCE", "formula").strip().lower()

__all__ = ["WEIGHT_SOURCE"]
