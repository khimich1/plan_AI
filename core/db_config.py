#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Единый источник путей к базам данных проекта.
Используется всеми модулями для доступа к pb.db и plita.db.
"""
from app.core.settings import get_settings

settings = get_settings()

PB_DB_PATH = settings.pb_db_path
PLITA_DB_PATH = settings.plita_db_path
