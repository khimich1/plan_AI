#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для инициализации менеджеров в базе данных plita.db

Простыми словами:
- Добавляет менеджеров из вашей таблицы в базу данных
- Если менеджер уже есть (по email), пропускает его
- Можно запускать несколько раз — ничего не сломается

Использование:
    python3 scripts/init_managers.py
"""

import sys
import os

# Добавляем корневую папку проекта в путь
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJECT_ROOT)

# Импортируем напрямую, чтобы избежать проблем с зависимостями
sys.path.insert(0, os.path.join(PROJECT_ROOT, 'core'))
from kp_db import init_default_managers, get_all_managers

if __name__ == '__main__':
    print("=" * 50)
    print("ИНИЦИАЛИЗАЦИЯ МЕНЕДЖЕРОВ В БАЗЕ ДАННЫХ")
    print("=" * 50)
    print()
    
    # Добавляем менеджеров по умолчанию
    print("📝 Добавляю менеджеров...")
    added_count = init_default_managers()
    print()
    
    # Показываем всех менеджеров
    print("📋 Список всех менеджеров в базе:")
    managers = get_all_managers()
    
    if managers:
        for manager in managers:
            print(f"  • {manager['fio']}")
            print(f"    Телефон: {manager['contact_number']}")
            print(f"    Email: {manager['email']}")
            print()
    else:
        print("  (менеджеров пока нет)")
    
    print("=" * 50)
    print(f"✅ Готово! Всего менеджеров в базе: {len(managers)}")
    print("=" * 50)
