#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Managers persistence — slice of kp_db (A1 decomposition)."""

from __future__ import annotations

import json
import logging
import os
import sqlite3
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from core.kp_db_common import DEFAULT_DB, _connect


def add_manager(
    fio: str,
    contact_number: str,
    email: str,
    db_path: str = DEFAULT_DB
) -> int:
    """
    Добавляет нового менеджера в базу данных.
    
    Простыми словами:
    - Сохраняет информацию о менеджере (ФИО, телефон, email)
    - Email должен быть уникальным (нельзя добавить двух менеджеров с одинаковым email)
    - Возвращает ID созданного менеджера
    
    Аргументы:
        fio: полное имя менеджера (например: "Иванов Иван Иванович")
        contact_number: контактный номер телефона (например: "79621860029")
        email: email адрес (например: "ivanov@example.ru")
        db_path: путь к базе данных
    
    Возвращает:
        ID созданного менеджера или 0 при ошибке
    """
    conn = _connect(db_path)
    
    try:
        cur = conn.cursor()
        cur.execute('''
            INSERT INTO managers (fio, contact_number, email)
            VALUES (?, ?, ?)
        ''', (fio, contact_number, email))
        
        manager_id = cur.lastrowid
        conn.commit()
        print(f"[DB] ✅ Менеджер добавлен: {fio} (ID: {manager_id})")
        return manager_id
        
    except sqlite3.IntegrityError:
        print(f"[DB] ⚠️ Менеджер с email {email} уже существует")
        conn.rollback()
        return 0
    except Exception as e:
        print(f"[DB] ❌ Ошибка при добавлении менеджера: {e}")
        conn.rollback()
        return 0
    
    finally:
        conn.close()


def get_all_managers(db_path: str = DEFAULT_DB) -> List[Dict]:
    """
    Получает список всех менеджеров.
    
    Простыми словами:
    - Возвращает всех менеджеров из базы данных
    - Каждый менеджер представлен словарём с полями: id, fio, contact_number, email
    
    Возвращает:
        Список словарей с информацией о менеджерах
    """
    conn = _connect(db_path)
    
    try:
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute('SELECT * FROM managers ORDER BY fio')
        return [dict(row) for row in cur.fetchall()]
    
    finally:
        conn.close()


def get_manager_by_id(manager_id: int, db_path: str = DEFAULT_DB) -> Optional[Dict]:
    """
    Получает информацию о менеджере по ID.
    
    Простыми словами:
    - Ищет менеджера по его порядковому номеру
    - Возвращает информацию о нём или None, если не найден
    
    Аргументы:
        manager_id: порядковый номер менеджера
        db_path: путь к базе данных
    
    Возвращает:
        Словарь с информацией о менеджере или None
    """
    conn = _connect(db_path)
    
    try:
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute('SELECT * FROM managers WHERE id = ?', (manager_id,))
        row = cur.fetchone()
        return dict(row) if row else None
    
    finally:
        conn.close()


def get_manager_by_email(email: str, db_path: str = DEFAULT_DB) -> Optional[Dict]:
    """
    Получает информацию о менеджере по email.
    
    Простыми словами:
    - Ищет менеджера по его email адресу
    - Возвращает информацию о нём или None, если не найден
    
    Аргументы:
        email: email адрес менеджера
        db_path: путь к базе данных
    
    Возвращает:
        Словарь с информацией о менеджере или None
    """
    conn = _connect(db_path)
    
    try:
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute('SELECT * FROM managers WHERE email = ?', (email,))
        row = cur.fetchone()
        return dict(row) if row else None
    
    finally:
        conn.close()


def update_manager(
    manager_id: int,
    fio: str = None,
    contact_number: str = None,
    email: str = None,
    db_path: str = DEFAULT_DB
) -> bool:
    """
    Обновляет информацию о менеджере.
    
    Простыми словами:
    - Меняет данные менеджера (можно обновить только нужные поля)
    - Если передать None для какого-то поля, оно не изменится
    
    Аргументы:
        manager_id: порядковый номер менеджера
        fio: новое ФИО (опционально)
        contact_number: новый контактный номер (опционально)
        email: новый email (опционально)
        db_path: путь к базе данных
    
    Возвращает:
        True если успешно, False если менеджер не найден
    """
    conn = _connect(db_path)
    
    try:
        cur = conn.cursor()
        
        # Формируем список обновлений
        updates = []
        values = []
        
        if fio is not None:
            updates.append('fio = ?')
            values.append(fio)
        if contact_number is not None:
            updates.append('contact_number = ?')
            values.append(contact_number)
        if email is not None:
            updates.append('email = ?')
            values.append(email)
        
        if not updates:
            return False
        
        values.append(manager_id)
        query = f'UPDATE managers SET {", ".join(updates)} WHERE id = ?'
        
        cur.execute(query, values)
        conn.commit()
        
        if cur.rowcount > 0:
            print(f"[DB] ✅ Менеджер #{manager_id} обновлён")
            return True
        else:
            print(f"[DB] ⚠️ Менеджер #{manager_id} не найден")
            return False
    
    except sqlite3.IntegrityError:
        print(f"[DB] ⚠️ Менеджер с таким email уже существует")
        conn.rollback()
        return False
    except Exception as e:
        print(f"[DB] ❌ Ошибка при обновлении менеджера: {e}")
        conn.rollback()
        return False
    
    finally:
        conn.close()


def delete_manager(manager_id: int, db_path: str = DEFAULT_DB) -> bool:
    """
    Удаляет менеджера из базы данных.
    
    Простыми словами:
    - Удаляет менеджера по его порядковому номеру
    - Это необратимая операция
    
    Аргументы:
        manager_id: порядковый номер менеджера для удаления
        db_path: путь к базе данных
    
    Возвращает:
        True если менеджер был найден и удалён, False если не найден
    """
    conn = _connect(db_path)
    
    try:
        cur = conn.cursor()
        cur.execute('DELETE FROM managers WHERE id = ?', (manager_id,))
        conn.commit()
        
        if cur.rowcount > 0:
            print(f"[DB] ✅ Менеджер #{manager_id} удалён")
            return True
        else:
            print(f"[DB] ⚠️ Менеджер #{manager_id} не найден")
            return False
    
    except Exception as e:
        print(f"[DB] ❌ Ошибка при удалении менеджера: {e}")
        conn.rollback()
        return False
    
    finally:
        conn.close()


def _managers_seed_path(seed_path: Optional[str] = None):
    """Resolve path to managers seed JSON (env MANAGERS_SEED_PATH overrides default)."""
    from pathlib import Path

    from core.config.settings import PROJECT_ROOT

    if seed_path:
        return Path(seed_path)
    env_path = os.environ.get("MANAGERS_SEED_PATH")
    if env_path:
        return Path(env_path)
    return PROJECT_ROOT / "data" / "managers_seed.json"


def _load_managers_seed(seed_path: Optional[str] = None) -> List[Tuple[str, str, str]]:
    """Load manager rows from JSON: [{fio, contact_number, email}, ...]."""
    import json
    import logging

    path = _managers_seed_path(seed_path)
    if not path.is_file():
        logging.getLogger(__name__).warning(
            "Managers seed file not found: %s (skipped)",
            path,
        )
        return []

    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        logging.getLogger(__name__).warning(
            "Failed to read managers seed %s: %s",
            path,
            exc,
        )
        return []

    if not isinstance(raw, list):
        logging.getLogger(__name__).warning("Managers seed must be a JSON array: %s", path)
        return []

    rows: List[Tuple[str, str, str]] = []
    for item in raw:
        if not isinstance(item, dict):
            continue
        fio = str(item.get("fio", "")).strip()
        contact = str(item.get("contact_number", "")).strip()
        email = str(item.get("email", "")).strip()
        if fio and email:
            rows.append((fio, contact, email))
    return rows


def init_default_managers(
    db_path: str = DEFAULT_DB,
    seed_path: Optional[str] = None,
) -> int:
    """
    Добавляет менеджеров из JSON seed-файла в базу данных.

    Источник: ``data/managers_seed.json`` или ``MANAGERS_SEED_PATH`` / аргумент
    ``seed_path``. Если файл отсутствует — ничего не добавляется.

    Возвращает:
        Количество успешно добавленных менеджеров
    """
    default_managers = _load_managers_seed(seed_path)

    added_count = 0

    for fio, contact_number, email in default_managers:
        manager_id = add_manager(fio, contact_number, email, db_path)
        if manager_id > 0:
            added_count += 1

    print(f"[DB] ✅ Добавлено менеджеров: {added_count} из {len(default_managers)}")
    return added_count
