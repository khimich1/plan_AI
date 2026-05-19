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

_COMMERCIAL_OFFER_LOGO_NAME = "ЖБЛСТАРТ.png"
_COMMERCIAL_OFFER_LOGO_CANDIDATES = (
    BASE_DIR / "банк знаний" / _COMMERCIAL_OFFER_LOGO_NAME,
    BASE_DIR / "docker" / "assets" / _COMMERCIAL_OFFER_LOGO_NAME,
)


def resolve_commercial_offer_logo_path() -> Path | None:
    """Плашка КП (логотип): локально — «банк знаний», в Docker — docker/assets/."""
    env_path = (os.environ.get("COMMERCIAL_OFFER_LOGO_PATH") or "").strip()
    if env_path:
        p = Path(env_path)
        if p.is_file():
            return p
    for candidate in _COMMERCIAL_OFFER_LOGO_CANDIDATES:
        if candidate.is_file():
            return candidate
    return None

# БД цен для ILP/оптимизации: локально pb.db в корне; в Docker задаётся PB_DB_PATH (тот же файл, что KPI/профиль).
_env_price_db = (os.environ.get("PRICE_DB_PATH") or os.environ.get("PB_DB_PATH") or "").strip()
PRICE_DB_PATH = Path(_env_price_db) if _env_price_db else BASE_DIR / "pb.db"

__all__ = [
    "BASE_DIR",
    "PROJECT_ROOT",
    "PRICE_XLSX_PATH",
    "CUTS_DOCX_PATH",
    "PRICE_DB_PATH",
    "resolve_commercial_offer_logo_path",
]
