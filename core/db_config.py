#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Единый источник путей к базам данных проекта.
Используется всеми модулями для доступа к pb.db и plita.db.
"""
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent

PB_DB_PATH = PROJECT_ROOT / "pb.db"
PLITA_DB_PATH = PROJECT_ROOT / "plita.db"
