#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Резолв путей к артефактам проекта (immutable после импорта)."""

from __future__ import annotations

import os
from pathlib import Path

# Корень репозитория (на уровень выше core/)
PROJECT_ROOT = Path(__file__).resolve().parent.parent
BASE_DIR = PROJECT_ROOT

PRICE_XLSX_PATH = BASE_DIR / "банк знаний" / "Новые цены для прайса с 19.08.24.xlsx"
CUTS_DOCX_PATH = BASE_DIR / "банк знаний" / "Письмо Цены с 29.05.2024 цены на резы.docx"

# БД цен для ILP/оптимизации: локально pb.db в корне; в Docker задаётся PB_DB_PATH (тот же файл, что KPI/профиль).
_env_price_db = (os.environ.get("PRICE_DB_PATH") or os.environ.get("PB_DB_PATH") or "").strip()
PRICE_DB_PATH = Path(_env_price_db) if _env_price_db else BASE_DIR / "pb.db"

__all__ = [
    "BASE_DIR",
    "PROJECT_ROOT",
    "PRICE_XLSX_PATH",
    "CUTS_DOCX_PATH",
    "PRICE_DB_PATH",
]
