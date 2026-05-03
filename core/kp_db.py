#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
База данных для хранения коммерческих предложений (КП) в формате XLSX

База данных: plita.db

Структура таблиц:
- KP_offers: основная информация о КП
- kp_plates: позиции (плиты) в каждом КП
- kp_files: файлы XLSX (хранит сам файл как BLOB и путь)
- kp_meta: метаданные (статус КП)
- completed_plates: выполненные плиты (перенесены из kp_plates после завершения дня)
- plate_rests: остатки от резки плит
- managers: менеджеры (ФИО, контактный номер, email)
"""

import os
import sqlite3
from datetime import datetime
from typing import List, Dict, Optional, Tuple, TypedDict
from core.debug_paths import get_debug_log_path

# Путь к базе данных (в корне проекта)
DEFAULT_DB = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'plita.db')
_DEBUG_SESSION_LOG = get_debug_log_path("debug-d7e22e.log")
_DEBUG_NOMENCLATURE_LOG = get_debug_log_path("debug-00f316.log")
_DEBUG_LOG = get_debug_log_path("debug.log")
_DEBUG_AGENT_LOG = get_debug_log_path("debug-ebb546.log")
_DEBUG_LOG_A9176E = get_debug_log_path("debug-a9176e.log")
_DEBUG_LOG_B59370 = get_debug_log_path("debug-b59370.log")
_DEBUG_LOG_8E9428 = get_debug_log_path("debug-8e9428.log")


def _debug_session_write(run_id: str, hypothesis_id: str, location: str, message: str, data: Dict) -> None:
    """Пишет NDJSON в debug-d7e22e.log для Debug Mode."""
    try:
        import json
        line = json.dumps({
            "sessionId": "d7e22e",
            "runId": run_id,
            "hypothesisId": hypothesis_id,
            "location": location,
            "message": message,
            "data": data,
            "timestamp": int(__import__("time").time() * 1000),
        }, ensure_ascii=False) + "\n"
        with open(_DEBUG_SESSION_LOG, "a", encoding="utf-8") as f:
            f.write(line)
    except Exception:
        pass


def _connect(db_path: str) -> sqlite3.Connection:
    """
    Безопасное подключение к SQLite (plita.db).

    - WAL уменьшает риск повреждения базы при сбоях/перезапусках.
    - foreign_keys нужен для корректной работы связей между таблицами.
    """
    conn = sqlite3.connect(db_path)
    conn.execute('PRAGMA journal_mode=WAL')
    conn.execute('PRAGMA foreign_keys = ON')
    return conn


def _audit_append(
    cur: sqlite3.Cursor,
    *,
    plate_id: Optional[int],
    kp_id: int,
    plate_name: Optional[str],
    plan_id: Optional[str],
    day_number: Optional[int],
    from_status: Optional[str],
    to_status: str,
    qty: int,
    reason: str,
    actor: Optional[str],
) -> None:
    """Добавляет запись в ``plate_status_log`` через переданный курсор.

    Локальный helper в core/, чтобы избежать импорта app/ из core/ слоя.
    Логически дублирует :class:`app.repositories.plate_audit_repository.PlateAuditRepository`,
    но репозиторий — это публичный API для сервисов, а здесь — внутренняя
    точка, вызываемая из :func:`mark_plates_as_planned`,
    :func:`return_plates_to_production`, :func:`move_plates_to_completed`.

    Принимает уже открытый ``cur`` — каллер сам отвечает за commit/rollback.
    """
    cur.execute(
        """
        INSERT INTO plate_status_log (
            plate_id, kp_id, plate_name, plan_id, day_number,
            from_status, to_status, qty, reason, actor
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            plate_id,
            int(kp_id),
            plate_name,
            plan_id,
            day_number,
            from_status,
            to_status,
            int(qty),
            reason,
            actor,
        ),
    )


def init_schema(db_path: str = DEFAULT_DB) -> None:
    """
    Создаёт таблицы в базе данных, если их ещё нет.
    
    Простыми словами:
    - Проверяет, есть ли таблицы в БД
    - Если нет — создаёт их с нужными колонками
    - Это как создать пустую таблицу Excel с заголовками
    """
    conn = _connect(db_path)
    try:
        cur = conn.cursor()
        
        # Таблица 1: KP_offers - Основная информация о КП
        # PRIMARY KEY - порядковый номер КП (auto increment)
        cur.execute('''
            CREATE TABLE IF NOT EXISTS KP_offers (
                kp_id INTEGER PRIMARY KEY AUTOINCREMENT,
                creation_date TEXT NOT NULL,
                customer_name TEXT,
                manager_name TEXT,
                discount_percent REAL DEFAULT 0,
                subtotal REAL,
                vat_amount REAL,
                total_amount REAL,
                delivery_conditions TEXT,
                payment_conditions TEXT,
                execution_terms TEXT
            )
        ''')
        
        # Таблица 2: kp_plates - Позиции (плиты) в каждом КП
        # kp_id - связь с KP_offers
        # status - статус плиты: "в производстве" (доступна для планирования) или "в плане" (уже добавлена в план)
        # plan_id - ID плана, в который добавлена плита (для связи и отката при удалении плана)
        # nomenclature_id - уникальный идентификатор из prays_plity
        cur.execute('''
            CREATE TABLE IF NOT EXISTS kp_plates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kp_id INTEGER NOT NULL,
                position_number INTEGER,
                plate_name TEXT NOT NULL,
                length_m REAL,
                width_m REAL,
                load_class INTEGER,
                qty INTEGER NOT NULL,
                unit_weight REAL,
                total_weight REAL,
                discounted_price REAL,
                status TEXT DEFAULT 'в производстве',
                plan_id TEXT,
                length_dm_raw TEXT,
                nomenclature_id TEXT,
                FOREIGN KEY (kp_id) REFERENCES KP_offers(kp_id) ON DELETE CASCADE
            )
        ''')
        
        # Миграция: колонка unit_price для пересчёта скидки и генерации документов из архива
        try:
            cur.execute("ALTER TABLE kp_plates ADD COLUMN unit_price REAL")
            conn.commit()
        except sqlite3.OperationalError:
            pass  # колонка уже существует
        
        # Таблица 3: kp_files - Файлы XLSX
        # kp_id - связь с KP_offers
        # xlsx_file - сам файл как BLOB (двоичные данные)
        # file_path - путь к файлу
        cur.execute('''
            CREATE TABLE IF NOT EXISTS kp_files (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kp_id INTEGER NOT NULL,
                xlsx_file BLOB,
                file_path TEXT,
                FOREIGN KEY (kp_id) REFERENCES KP_offers(kp_id) ON DELETE CASCADE,
                UNIQUE(kp_id)
            )
        ''')
        
        # Таблица 4: kp_meta - Метаданные
        # kp_id - связь с KP_offers
        # status - статус КП: выполнено, отклонено, в работе, в ожидании
        cur.execute('''
            CREATE TABLE IF NOT EXISTS kp_meta (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kp_id INTEGER NOT NULL,
                status TEXT DEFAULT 'в работе',
                FOREIGN KEY (kp_id) REFERENCES KP_offers(kp_id) ON DELETE CASCADE,
                UNIQUE(kp_id)
            )
        ''')
        
        # Таблица 5: completed_plates - Выполненные плиты
        # Сюда переносятся плиты после завершения дня производства
        cur.execute('''
            CREATE TABLE IF NOT EXISTS completed_plates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kp_id INTEGER NOT NULL,
                plate_name TEXT NOT NULL,
                length_m REAL,
                width_m REAL,
                load_class INTEGER,
                qty INTEGER NOT NULL,
                completed_date TEXT NOT NULL,
                production_day INTEGER,
                nomenclature_id TEXT,
                FOREIGN KEY (kp_id) REFERENCES KP_offers(kp_id) ON DELETE CASCADE
            )
        ''')
        
        # Таблица 6: plate_rests - Остатки от резки плит
        # Хранит информацию об остатках, которые образуются при продольном резе
        # Статусы: available (доступен), used (использован), completed (выполнен), discarded (списан)
        cur.execute('''
            CREATE TABLE IF NOT EXISTS plate_rests (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kp_id INTEGER NOT NULL,
                source_plate_name TEXT NOT NULL,
                rest_width_mm INTEGER NOT NULL,
                length_m REAL NOT NULL,
                qty INTEGER NOT NULL DEFAULT 1,
                status TEXT DEFAULT 'available',
                created_date TEXT NOT NULL,
                used_date TEXT,
                production_day INTEGER,
                FOREIGN KEY (kp_id) REFERENCES KP_offers(kp_id) ON DELETE CASCADE
            )
        ''')
        
        # Таблица 7: managers - Менеджеры
        # Хранит информацию о менеджерах: ФИО, контактный номер, email
        cur.execute('''
            CREATE TABLE IF NOT EXISTS managers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fio TEXT NOT NULL,
                contact_number TEXT NOT NULL,
                email TEXT NOT NULL,
                UNIQUE(email)
            )
        ''')

        # Таблица 8: plate_status_log — журнал переходов статусов плит.
        # Каждая запись — это отдельный переход (planned / completed / rejected /
        # plan_rollback). Используется для диагностики «куда делась плита» и
        # независима от kp_plates: даже если строка в kp_plates удалена, история
        # переходов остаётся.
        cur.execute('''
            CREATE TABLE IF NOT EXISTS plate_status_log (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                plate_id     INTEGER,
                kp_id        INTEGER NOT NULL,
                plate_name   TEXT,
                plan_id      TEXT,
                day_number   INTEGER,
                from_status  TEXT,
                to_status    TEXT NOT NULL,
                qty          INTEGER NOT NULL,
                reason       TEXT,
                actor        TEXT,
                created_at   TEXT NOT NULL DEFAULT (datetime('now'))
            )
        ''')

        # Создаём индексы для быстрого поиска
        # Это как закладки в книге — помогают быстро найти нужную информацию
        cur.execute('CREATE INDEX IF NOT EXISTS idx_kp_id_plates ON kp_plates(kp_id)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_kp_id_files ON kp_files(kp_id)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_kp_id_meta ON kp_meta(kp_id)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_meta_status ON kp_meta(status)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_completed_kp_id ON completed_plates(kp_id)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_completed_date ON completed_plates(completed_date)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_rests_kp_id ON plate_rests(kp_id)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_rests_status ON plate_rests(status)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_managers_email ON managers(email)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_status_log_kp ON plate_status_log(kp_id)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_status_log_plan ON plate_status_log(plan_id)')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_status_log_status ON plate_status_log(to_status)')
        # Индексы для status и plan_id создаются в блоке миграции ниже
        
        # === МИГРАЦИЯ: Добавляем колонки status и plan_id если их нет ===
        # Это нужно для существующих баз данных, где таблица уже создана без этих полей
        cur.execute("PRAGMA table_info(kp_plates)")
        columns = [col[1] for col in cur.fetchall()]
        
        if 'status' not in columns:
            print("[DB] Миграция: добавляем колонку status в kp_plates...")
            cur.execute("ALTER TABLE kp_plates ADD COLUMN status TEXT DEFAULT 'в производстве'")
            # Устанавливаем статус для всех существующих записей
            cur.execute("UPDATE kp_plates SET status = 'в производстве' WHERE status IS NULL")
            print("[DB] ✅ Колонка status добавлена")
        
        # Создаём индекс для status (после того, как колонка точно существует)
        cur.execute('CREATE INDEX IF NOT EXISTS idx_plates_status ON kp_plates(status)')
        
        if 'plan_id' not in columns:
            print("[DB] Миграция: добавляем колонку plan_id в kp_plates...")
            cur.execute("ALTER TABLE kp_plates ADD COLUMN plan_id TEXT")
            print("[DB] ✅ Колонка plan_id добавлена")
        
        # Создаём индекс для plan_id (после того, как колонка точно существует)
        cur.execute('CREATE INDEX IF NOT EXISTS idx_plates_plan_id ON kp_plates(plan_id)')

        # P5: day_number — день производства, в который попала плита.
        # Записывается mark_plates_as_planned. Нужен, чтобы day_view мог читать
        # plates_info напрямую из БД и держать инвариант с complete_day.
        if 'day_number' not in columns:
            print("[DB] Миграция: добавляем колонку day_number в kp_plates...")
            cur.execute("ALTER TABLE kp_plates ADD COLUMN day_number INTEGER")
            print("[DB] ✅ Колонка day_number добавлена")

        cur.execute(
            'CREATE INDEX IF NOT EXISTS idx_plates_plan_day '
            'ON kp_plates(plan_id, day_number)'
        )
        
        if 'length_dm_raw' not in columns:
            print("[DB] Миграция: добавляем колонку length_dm_raw в kp_plates...")
            cur.execute("ALTER TABLE kp_plates ADD COLUMN length_dm_raw TEXT")
            print("[DB] ✅ Колонка length_dm_raw добавлена")
        
        # Обратная засылка length_dm_raw из plate_name для существующих строк (один раз после ADD COLUMN)
        cur.execute("SELECT id, plate_name, length_m FROM kp_plates WHERE length_dm_raw IS NULL OR length_dm_raw = ''")
        rows = cur.fetchall()
        if rows:
            from core import config_and_data as cfg
            for row in rows:
                rid, plate_name, length_m = row
                raw = cfg.extract_length_dm_raw_from_plate_name(plate_name or "")
                if raw:
                    cur.execute("UPDATE kp_plates SET length_dm_raw = ? WHERE id = ?", (raw, rid))
                elif length_m is not None:
                    # Fallback: из length_m (5.98 -> "59,8", 6.12 -> "61,2")
                    dm = float(length_m) * 10
                    if abs(dm - round(dm)) < 0.01:
                        raw_fb = str(int(round(dm)))
                    else:
                        raw_fb = f"{dm:.1f}".rstrip('0').rstrip('.').replace('.', ',')
                    cur.execute("UPDATE kp_plates SET length_dm_raw = ? WHERE id = ?", (raw_fb, rid))
            if rows:
                print(f"[DB] ✅ Обратная засылка length_dm_raw: обновлено {len(rows)} строк")
        
        # === МИГРАЦИЯ: Добавляем колонку status в kp_meta если её нет ===
        # Это нужно для существующих баз данных, где таблица уже создана без этого поля
        cur.execute("PRAGMA table_info(kp_meta)")
        meta_columns = [col[1] for col in cur.fetchall()]
        
        if 'status' not in meta_columns:
            print("[DB] Миграция: добавляем колонку status в kp_meta...")
            cur.execute("ALTER TABLE kp_meta ADD COLUMN status TEXT DEFAULT 'в работе'")
            # Устанавливаем статус для всех существующих записей
            cur.execute("UPDATE kp_meta SET status = 'в работе' WHERE status IS NULL")
            print("[DB] ✅ Колонка status добавлена в kp_meta")

        # === МИГРАЦИЯ: Добавляем nomenclature_id ===
        if 'nomenclature_id' not in columns:
            print("[DB] Миграция: добавляем колонку nomenclature_id в kp_plates...")
            cur.execute("ALTER TABLE kp_plates ADD COLUMN nomenclature_id TEXT")
            print("[DB] Колонка nomenclature_id добавлена в kp_plates")

        cur.execute("PRAGMA table_info(completed_plates)")
        cp_columns = [col[1] for col in cur.fetchall()]
        if 'nomenclature_id' not in cp_columns:
            print("[DB] Миграция: добавляем колонку nomenclature_id в completed_plates...")
            cur.execute("ALTER TABLE completed_plates ADD COLUMN nomenclature_id TEXT")
            print("[DB] Колонка nomenclature_id добавлена в completed_plates")
        
        conn.commit()
    finally:
        conn.close()


_PB_DB_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'pb.db')


def _extract_length_dm_val(name: str) -> Optional[float]:
    """Возвращает числовое значение длины в дм из марки плиты или None."""
    import re as _re
    m = _re.search(r'П[БК]\s*([\d,\.]+)\s*-', str(name or ''))
    if not m:
        return None
    try:
        return float(m.group(1).replace(',', '.'))
    except ValueError:
        return None


def lookup_nomenclature_by_plate_name(
    plate_name: str,
    pb_cur: sqlite3.Cursor,
) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """Ищет запись в prays_plity по имени плиты.

    Возвращает (canonical_name, nomenclature_id, match_type).
    match_type: "exact" | "like" | None (не найдено).
    """
    # Точное совпадение (без учёта регистра)
    pb_cur.execute(
        'SELECT "Уникальный идентификатор (Номенклатура)", "Товар" '
        'FROM prays_plity WHERE "Товар" = ? COLLATE NOCASE',
        (plate_name,),
    )
    row = pb_cur.fetchone()
    if row:
        return row[1], row[0], "exact"

    # Варианты формата: ширина -N-→-N,0-, длина 45→45,0 и комбинации
    from core.config_and_data import plate_name_to_prays_variants
    for prays_variant in plate_name_to_prays_variants(plate_name):
        pb_cur.execute(
            'SELECT "Уникальный идентификатор (Номенклатура)", "Товар" '
            'FROM prays_plity WHERE "Товар" = ? COLLATE NOCASE',
            (prays_variant,),
        )
        row = pb_cur.fetchone()
        if row:
            return row[1], row[0], "exact_prays_variant"

    # Частичное совпадение: убираем «Плиты »/«Плита » и ищем LIKE.
    # Защита от ложных совпадений: числовое значение длины в найденном canonical_name
    # должно совпадать с длиной в plate_name (допуск 0.05 дм). Это предотвращает
    # подстановку «57-12-8п» вместо «57,1-12-8п» из-за частичного совпадения LIKE.
    req_len_val = _extract_length_dm_val(plate_name)
    normalized = plate_name.replace('Плиты ', '').replace('Плита ', '')
    pb_cur.execute(
        'SELECT "Уникальный идентификатор (Номенклатура)", "Товар" '
        'FROM prays_plity WHERE "Товар" LIKE ?',
        (f'%{normalized}%',),
    )
    row = pb_cur.fetchone()
    if row:
        can_len_val = _extract_length_dm_val(row[1])
        if req_len_val is not None and can_len_val is not None:
            if abs(req_len_val - can_len_val) > 0.05:
                # Длины не совпадают — ложное LIKE-совпадение, игнорируем
                row = None
    if row:
        return row[1], row[0], "like"

    # #region agent log (nomenclature not found — stage: after exact, variants, LIKE)
    try:
        _debug_log = _DEBUG_NOMENCLATURE_LOG
        with open(_debug_log, 'a', encoding='utf-8') as _f:
            _f.write(__import__('json').dumps({"sessionId": "00f316", "hypothesisId": "nomenclature_stage", "location": "kp_db:lookup_nomenclature_by_plate_name", "message": "nomenclature not found after exact, variants, LIKE", "data": {"plate_name": (plate_name or "")[:120], "stage": "not_found"}, "timestamp": __import__("time").time() * 1000}, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion
    return None, None, None


def fill_plate_nomenclature_cache() -> None:
    """Заполняет PLATE_NOMENCLATURE_CACHE из prays_plity для всех позиций в PLATE_LOAD_DETAILS.

    Для каждого ключа (length, width, load_code) формирует имя через make_plate_name
    с length_dm_raw из PLATE_LENGTH_DM_RAW, ищет его в prays_plity и кладёт результат
    в PLATE_NOMENCLATURE_CACHE. При отсутствии pb.db — безопасно выходит.
    """
    import logging as _logging
    _log = _logging.getLogger(__name__)

    from core import config_and_data as cfg

    if not os.path.exists(_PB_DB_PATH):
        _log.debug("fill_plate_nomenclature_cache: pb.db не найден, кэш не заполняется")
        return

    pb_conn = sqlite3.connect(_PB_DB_PATH)
    try:
        pb_cur = pb_conn.cursor()
        for key, _qty in cfg.PLATE_LOAD_DETAILS.items():
            # Не обновляем ключи, которые уже в кэше (повторный вызов после apply_to_globals)
            if key in cfg.PLATE_NOMENCLATURE_CACHE:
                continue
            length_m, width_m, load_code = key[0], key[1], key[2]
            length_dm_raw = key[3] if len(key) > 3 else cfg.PLATE_LENGTH_DM_RAW.get(key, "")
            plate_name = cfg.make_plate_name(length_m, width_m, load_code=load_code, length_dm_raw=length_dm_raw)
            canonical_name, nomenclature_id, _ = lookup_nomenclature_by_plate_name(plate_name, pb_cur)
            cfg.PLATE_NOMENCLATURE_CACHE[key] = {
                "canonical_name": canonical_name,
                "nomenclature_id": nomenclature_id,
            }
            # #region agent log (a9176e: 57/57,1 — fill_plate_nomenclature_cache)
            if 5.69 <= length_m <= 5.73:
                try:
                    _log_path = _DEBUG_LOG_A9176E
                    with open(_log_path, 'a', encoding='utf-8') as _f:
                        _f.write(__import__('json').dumps({"sessionId": "a9176e", "hypothesisId": "H5", "location": "kp_db:fill_plate_nomenclature_cache", "message": "57/57,1 cache fill", "data": {"key": [length_m, width_m, load_code], "length_dm_raw": length_dm_raw, "plate_name": (plate_name or "")[:60], "canonical_name": (canonical_name or "")[:60] if canonical_name else None}, "timestamp": __import__("time").time() * 1000}, ensure_ascii=False) + "\n")
                except Exception:
                    pass
            # #endregion
            _log.debug(
                f"Кэш: {plate_name!r} → canonical={canonical_name!r}, id={nomenclature_id!r}"
            )
    except Exception as e:
        _log.warning(f"fill_plate_nomenclature_cache: ошибка: {e}")
    finally:
        pb_conn.close()


def enrich_order_data_with_nomenclature(order_data: List[Dict]) -> List[Dict]:
    """
    Обогащает данные заказа точными названиями плит и уникальными идентификаторами (nomenclature_id)
    из базы данных prays_plity (pb.db).

    Элементы, у которых nomenclature_id уже задан (заполнен из PLATE_NOMENCLATURE_CACHE
    на этапе сборки позиций), пропускаются — их данные не перезаписываются.
    Остальные элементы обогащаются как fallback (например, позиции из плана оптимизации).
    """
    if not os.path.exists(_PB_DB_PATH):
        print(f"[DB] ⚠️ Файл pb.db не найден по пути: {_PB_DB_PATH}")
        return order_data

    pb_conn = sqlite3.connect(_PB_DB_PATH)
    try:
        pb_cur = pb_conn.cursor()

        for item in order_data:
            # Пропускаем элементы, уже обогащённые через кэш
            if item.get('nomenclature_id') is not None:
                # #region agent log (00f316: плита уже с id из кэша)
                try:
                    _debug_log = _DEBUG_NOMENCLATURE_LOG
                    with open(_debug_log, 'a', encoding='utf-8') as _f:
                        _f.write(__import__('json').dumps({"sessionId": "00f316", "hypothesisId": "nomenclature_check", "location": "kp_db:enrich_order_data_with_nomenclature", "message": "plate skipped, has nomenclature_id from cache", "data": {"plate_name": (item.get('name', '') or '')[:120], "has_nomenclature_id": True, "match_type": "from_cache"}, "timestamp": __import__("time").time() * 1000}, ensure_ascii=False) + "\n")
                except Exception:
                    pass
                # #endregion
                # #region agent log
                try:
                    _log_path = _DEBUG_LOG_B59370
                    with open(_log_path, 'a', encoding='utf-8') as _f:
                        _f.write(__import__('json').dumps({"sessionId": "b59370", "hypothesisId": "H_prays", "location": "kp_db:enrich_skipped", "message": "item had nomenclature_id from cache", "data": {"name": (item.get('name', '') or '')[:60]}, "timestamp": __import__("time").time() * 1000}, ensure_ascii=False) + "\n")
                except Exception:
                    pass
                # #endregion
                continue

            original_name = item.get('name', '')
            canonical_name, nomenclature_id, match_type = lookup_nomenclature_by_plate_name(original_name, pb_cur)
            if canonical_name is not None:
                item['name'] = canonical_name
                item['nomenclature_id'] = nomenclature_id
            else:
                item['nomenclature_id'] = None
            # #region agent log (a9176e: 57/57,1 — enrich подстановка)
            if ('57,1' in (original_name or '') or ('57-12' in (original_name or '') and '57,' not in (original_name or ''))):
                try:
                    _log_path = _DEBUG_LOG_A9176E
                    with open(_log_path, 'a', encoding='utf-8') as _f:
                        _f.write(__import__('json').dumps({"sessionId": "a9176e", "hypothesisId": "H4", "location": "kp_db:enrich_order_data_with_nomenclature", "message": "57/57,1 enrich lookup", "data": {"original_name": (original_name or "")[:80], "canonical_name": (canonical_name or "")[:80] if canonical_name else None, "match_type": match_type, "replaced": canonical_name is not None}, "timestamp": __import__("time").time() * 1000}, ensure_ascii=False) + "\n")
                except Exception:
                    pass
            # #endregion
            # #region agent log (00f316: каждая плита — есть ли nomenclature_id и этап совпадения)
            try:
                _debug_log = _DEBUG_NOMENCLATURE_LOG
                with open(_debug_log, 'a', encoding='utf-8') as _f:
                    _f.write(__import__('json').dumps({"sessionId": "00f316", "hypothesisId": "nomenclature_check", "location": "kp_db:enrich_order_data_with_nomenclature", "message": "plate nomenclature result", "data": {"plate_name": (original_name or "")[:120], "has_nomenclature_id": nomenclature_id is not None, "match_type": match_type}, "timestamp": __import__("time").time() * 1000}, ensure_ascii=False) + "\n")
            except Exception:
                pass
            # #endregion
            # #region agent log (плиты без номенклатуры в prays: 61,8-5 / 45-7 / 37,9-9)
            _no_nom_substrings = ('61,8-5', '45-7', '37,9-9')
            if any(sub in (original_name or '') for sub in _no_nom_substrings):
                try:
                    _log_path = _DEBUG_LOG_8E9428
                    with open(_log_path, 'a', encoding='utf-8') as _f:
                        _f.write(__import__('json').dumps({"sessionId": "8e9428", "hypothesisId": "H_no_nomenclature", "location": "kp_db:enrich_order_data_with_nomenclature", "message": "lookup prays_plity for 61,8-5 / 45-7 / 37,9-9", "data": {"original_name": (original_name or "")[:80], "canonical_name": (canonical_name[:80] if canonical_name else None), "nomenclature_id": nomenclature_id, "match_type": match_type}, "timestamp": __import__("time").time() * 1000}, ensure_ascii=False) + "\n")
                except Exception:
                    pass
            # #endregion
            # #region agent log
            try:
                _log_path = _DEBUG_LOG_B59370
                with open(_log_path, 'a', encoding='utf-8') as _f:
                    _f.write(__import__('json').dumps({"sessionId": "b59370", "hypothesisId": "H_prays", "location": "kp_db:enrich_lookup", "message": "prays_plity lookup result", "data": {"original_name": original_name[:60], "canonical_name": (canonical_name[:60] if canonical_name else None), "match_type": match_type}, "timestamp": __import__("time").time() * 1000}, ensure_ascii=False) + "\n")
            except Exception:
                pass
            # #endregion

    except Exception as e:
        print(f"[DB] ❌ Ошибка при обогащении order_data номенклатурами: {e}")
    finally:
        pb_conn.close()

    return order_data


def save_kp_to_db(
    creation_date: str,
    order_data: List[Dict],
    xlsx_file_path: Optional[str] = None,
    customer_name: Optional[str] = None,
    manager_name: Optional[str] = None,
    discount_percent: float = 0,
    delivery_conditions: Optional[str] = None,
    payment_conditions: Optional[str] = None,
    execution_terms: Optional[str] = None,
    status: str = 'в работе',
    db_path: str = DEFAULT_DB
) -> int:
    """
    Сохраняет КП в базу данных.
    
    Простыми словами:
    - Берёт всю информацию о КП
    - Записывает её в базу данных
    - Возвращает порядковый номер КП (kp_id)
    
    Аргументы:
        creation_date: дата создания КП (например: "01.01.2024" или "2024-01-01")
        order_data: список плит в заказе (из функции generate_commercial_offer_xlsx)
        xlsx_file_path: путь к файлу XLSX (если есть)
        customer_name: имя клиента
        manager_name: имя менеджера
        discount_percent: процент скидки (0-100)
        delivery_conditions: условия поставки
        payment_conditions: условия оплаты
        execution_terms: сроки выполнения
        status: статус КП (по умолчанию "в работе")
        db_path: путь к базе данных
    
    Возвращает:
        Порядковый номер КП (kp_id)
    """
    init_schema(db_path)
    
    # 🔥 ИСПРАВЛЕНИЕ: Используем ту же функцию расчета, что и в XLSX
    # Это гарантирует, что суммы в БД и в XLSX файле будут одинаковыми
    try:
        from core.commercial_offer_xlsx import calculate_total_cost
        totals = calculate_total_cost(order_data, discount_percent)
        subtotal = totals['subtotal']
        vat_amount = totals['vat_amount']
        total_amount = totals['total_with_vat']
    except ImportError:
        # Fallback: старая логика, если модуль не найден
        subtotal = 0.0
        for item in order_data:
            qty = item.get('qty', 0)
            unit_price = item.get('unit_price', 0.0)
            discounted_price = unit_price * (1 - discount_percent / 100)
            subtotal += discounted_price * qty
        
        vat_amount = round(subtotal * 0.22, 2)
        total_amount = round(subtotal + vat_amount, 2)
    
    conn = _connect(db_path)
    try:
        # Включаем поддержку FOREIGN KEY
        conn.execute('PRAGMA foreign_keys = ON')
        
        cur = conn.cursor()
        
        # Сохраняем основную информацию о КП
        cur.execute('''
            INSERT INTO KP_offers (
                creation_date, customer_name, manager_name, discount_percent,
                subtotal, vat_amount, total_amount,
                delivery_conditions, payment_conditions, execution_terms
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            creation_date, customer_name, manager_name, discount_percent,
            subtotal, vat_amount, total_amount,
            delivery_conditions, payment_conditions, execution_terms
        ))
        
        # Получаем порядковый номер созданного КП
        kp_id = cur.lastrowid
        
        # Сохраняем позиции (плиты)
        for idx, item in enumerate(order_data, start=1):
            qty = item.get('qty', 0)
            unit_price = item.get('unit_price', 0.0)
            discounted_price = unit_price * (1 - discount_percent / 100)
            weight = item.get('weight', 0.0)
            unit_weight = weight / qty if qty > 0 else 0.0
            
            # Точное название и nomenclature_id уже обогащены перед вызовом save_kp_to_db
            plate_name = item.get('name', '')
            nomenclature_id = item.get('nomenclature_id', None)
            
            # Длина из КП сохраняется как есть; length_dm_raw — исходная строка из марки для поиска при списании.
            cur.execute('''
                INSERT INTO kp_plates (
                    kp_id, position_number, plate_name,
                    length_m, width_m, load_class,
                    qty, unit_weight, total_weight, discounted_price, unit_price,
                    length_dm_raw, nomenclature_id
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                kp_id, idx, plate_name,
                item.get('length_m', 0), item.get('width_m', 0), item.get('load_class', 800),
                qty, unit_weight, weight, discounted_price, unit_price,
                item.get('length_dm_raw', '') or '', nomenclature_id
            ))
            
        # Сохраняем файл XLSX (если указан путь)
        if xlsx_file_path and os.path.exists(xlsx_file_path):
            with open(xlsx_file_path, 'rb') as f:
                xlsx_blob = f.read()
            
            cur.execute('''
                INSERT INTO kp_files (kp_id, xlsx_file, file_path)
                VALUES (?, ?, ?)
            ''', (kp_id, xlsx_blob, xlsx_file_path))
        elif xlsx_file_path:
            # Если путь указан, но файла нет, сохраняем только путь
            cur.execute('''
                INSERT INTO kp_files (kp_id, xlsx_file, file_path)
                VALUES (?, ?, ?)
            ''', (kp_id, None, xlsx_file_path))
        
        # Сохраняем метаданные (статус)
        cur.execute('''
            INSERT INTO kp_meta (kp_id, status)
            VALUES (?, ?)
        ''', (kp_id, status))
        
        conn.commit()
        return kp_id
    
    finally:
        conn.close()


def update_kp_discount(kp_id: int, new_discount: float, db_path: str = DEFAULT_DB) -> bool:
    """
    Обновляет процент скидки для КП, пересчитывает итоги и discounted_price по плитам.
    
    Простыми словами:
    - Берёт unit_price из kp_plates (или восстанавливает из discounted_price и текущей скидки)
    - Считает новые subtotal/vat/total через calculate_total_cost
    - Обновляет KP_offers и discounted_price (и при необходимости unit_price) в kp_plates
    """
    if not (0 <= new_discount <= 100):
        return False
    init_schema(db_path)
    try:
        from core.commercial_offer_xlsx import calculate_total_cost
    except ImportError:
        return False
    
    conn = _connect(db_path)
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()
        
        cur.execute('SELECT discount_percent FROM KP_offers WHERE kp_id = ?', (kp_id,))
        row = cur.fetchone()
        if not row:
            return False
        current_discount = row[0] or 0.0
        
        cur.execute(
            'SELECT id, plate_name, length_m, width_m, load_class, qty, unit_weight, total_weight, discounted_price, unit_price FROM kp_plates WHERE kp_id = ? ORDER BY position_number',
            (kp_id,)
        )
        plates = cur.fetchall()
        if not plates:
            return False
        
        order_data = []
        for p in plates:
            pid, plate_name, length_m, width_m, load_class, qty, unit_weight, total_weight, discounted_price, unit_price_col = p
            if unit_price_col is not None and unit_price_col > 0:
                unit_price = float(unit_price_col)
            else:
                # Восстанавливаем из discounted_price и текущей скидки
                factor = 1.0 - (current_discount / 100.0)
                if factor <= 0:
                    factor = 1.0
                unit_price = (discounted_price or 0) / factor
            weight = total_weight if total_weight is not None and total_weight > 0 else (unit_weight or 0) * (qty or 0)
            order_data.append({
                'name': plate_name or '',
                'length_m': length_m or 0,
                'width_m': width_m or 0,
                'qty': qty or 0,
                'load_class': load_class or 800,
                'unit_price': unit_price,
                'weight': weight,
            })
        
        totals = calculate_total_cost(order_data, new_discount)
        subtotal = totals['subtotal']
        vat_amount = totals['vat_amount']
        total_amount = totals['total_with_vat']
        
        cur.execute('''
            UPDATE KP_offers SET discount_percent = ?, subtotal = ?, vat_amount = ?, total_amount = ?
            WHERE kp_id = ?
        ''', (new_discount, subtotal, vat_amount, total_amount, kp_id))
        
        for p in plates:
            pid, plate_name, length_m, width_m, load_class, qty, unit_weight, total_weight, discounted_price, unit_price_col = p
            if unit_price_col is not None and unit_price_col > 0:
                up = float(unit_price_col)
            else:
                factor = 1.0 - (current_discount / 100.0)
                if factor <= 0:
                    factor = 1.0
                up = (discounted_price or 0) / factor
            new_disc_price = up * (1.0 - new_discount / 100.0)
            cur.execute(
                'UPDATE kp_plates SET discounted_price = ?, unit_price = ? WHERE id = ?',
                (new_disc_price, up, pid)
            )
        
        conn.commit()
        return True
    except Exception:
        return False
    finally:
        conn.close()


def get_kp_by_id(kp_id: int, db_path: str = DEFAULT_DB) -> Optional[Dict]:
    """
    Получает информацию о КП по порядковому номеру.
    
    Простыми словами:
    - Ищет КП по номеру в базе
    - Возвращает всю информацию о нём + список плит + файл + статус
    
    Возвращает:
        Словарь с информацией о КП или None, если не найдено
    """
    init_schema(db_path)
    conn = _connect(db_path)
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        
        # Получаем основную информацию
        cur.execute('SELECT * FROM KP_offers WHERE kp_id = ?', (kp_id,))
        row = cur.fetchone()
        if not row:
            return None
        
        kp_data = dict(row)
        
        # Получаем список плит
        cur.execute('''
            SELECT * FROM kp_plates 
            WHERE kp_id = ? 
            ORDER BY position_number
        ''', (kp_id,))
        plates = [dict(row) for row in cur.fetchall()]
        kp_data['plates'] = plates
        
        # Получаем информацию о файле
        cur.execute('SELECT * FROM kp_files WHERE kp_id = ?', (kp_id,))
        file_row = cur.fetchone()
        if file_row:
            file_data = dict(file_row)
            # BLOB не нужно конвертировать в строку, оставляем как bytes
            kp_data['file'] = file_data
        
        # Получаем статус из метаданных
        cur.execute('SELECT status FROM kp_meta WHERE kp_id = ?', (kp_id,))
        meta_row = cur.fetchone()
        if meta_row:
            kp_data['status'] = meta_row['status']
        
        return kp_data
    
    finally:
        conn.close()


def get_all_kp_by_status(status: str, db_path: str = DEFAULT_DB) -> List[Dict]:
    """
    Получает все КП с определённым статусом.
    
    Простыми словами:
    - Ищет все КП со статусом (например, "в работе")
    - Возвращает список таких КП
    
    Аргументы:
        status: статус для поиска ("выполнено", "отклонено", "в работе", "в ожидании")
    
    Возвращает:
        Список словарей с информацией о КП
    """
    init_schema(db_path)
    conn = _connect(db_path)
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        
        # Получаем ID всех КП с нужным статусом
        cur.execute('SELECT kp_id FROM kp_meta WHERE status = ?', (status,))
        kp_ids = [row['kp_id'] for row in cur.fetchall()]
        
        # Получаем информацию о каждом КП
        result = []
        for kp_id in kp_ids:
            kp_data = get_kp_by_id(kp_id, db_path)
            if kp_data:
                result.append(kp_data)
        
        return result
    
    finally:
        conn.close()


def update_kp_status(kp_id: int, new_status: str, db_path: str = DEFAULT_DB) -> bool:
    """
    Обновляет статус КП.
    
    Простыми словами:
    - Меняет статус КП (например, с "в работе" на "выполнено")
    
    Аргументы:
        kp_id: порядковый номер КП
        new_status: новый статус ("выполнено", "отклонено", "в работе", "в ожидании")
    
    Возвращает:
        True если успешно, False если КП не найдено
    """
    init_schema(db_path)
    conn = _connect(db_path)
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()
        cur.execute('''
            UPDATE kp_meta 
            SET status = ?
            WHERE kp_id = ?
        ''', (new_status, kp_id))
        
        conn.commit()
        return cur.rowcount > 0
    
    finally:
        conn.close()


def save_xlsx_file(kp_id: int, xlsx_file_path: str, db_path: str = DEFAULT_DB) -> bool:
    """
    Сохраняет файл XLSX в базу данных (обновляет или создаёт запись).
    
    Простыми словами:
    - Читает файл XLSX с диска
    - Сохраняет его в базу данных как двоичные данные
    - Также сохраняет путь к файлу
    
    Возвращает:
        True если успешно, False если ошибка
    """
    init_schema(db_path)
    
    if not os.path.exists(xlsx_file_path):
        return False
    
    conn = _connect(db_path)
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        with open(xlsx_file_path, 'rb') as f:
            xlsx_blob = f.read()
        
        cur = conn.cursor()
        # Проверяем, есть ли уже запись
        cur.execute('SELECT id FROM kp_files WHERE kp_id = ?', (kp_id,))
        exists = cur.fetchone()
        
        if exists:
            # Обновляем существующую запись
            cur.execute('''
                UPDATE kp_files 
                SET xlsx_file = ?, file_path = ?
                WHERE kp_id = ?
            ''', (xlsx_blob, xlsx_file_path, kp_id))
        else:
            # Создаём новую запись
            cur.execute('''
                INSERT INTO kp_files (kp_id, xlsx_file, file_path)
                VALUES (?, ?, ?)
            ''', (kp_id, xlsx_blob, xlsx_file_path))
        
        conn.commit()
        return True
    
    finally:
        conn.close()


def get_xlsx_file(kp_id: int, output_path: Optional[str] = None, db_path: str = DEFAULT_DB) -> Optional[bytes]:
    """
    Получает файл XLSX из базы данных.
    
    Простыми словами:
    - Извлекает файл XLSX из базы данных
    - Если указан output_path, сохраняет файл на диск
    - Иначе возвращает файл как bytes (двоичные данные)
    
    Аргументы:
        kp_id: порядковый номер КП
        output_path: путь для сохранения файла (опционально)
        db_path: путь к базе данных
    
    Возвращает:
        bytes (двоичные данные файла) или None если не найдено
    """
    init_schema(db_path)
    conn = _connect(db_path)
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()
        cur.execute('SELECT xlsx_file FROM kp_files WHERE kp_id = ?', (kp_id,))
        row = cur.fetchone()
        
        if not row or not row[0]:
            return None
        
        xlsx_data = row[0]
        
        # Если указан путь для сохранения, сохраняем файл
        if output_path:
            with open(output_path, 'wb') as f:
                f.write(xlsx_data)
        
        return xlsx_data
    
    finally:
        conn.close()


def delete_kp_by_id(kp_id: int, db_path: str = DEFAULT_DB) -> bool:
    """
    Удаляет КП из базы данных по порядковому номеру.
    
    Простыми словами:
    - Ищет КП по номеру
    - Удаляет его из базы (включая все плиты, файлы, метаданные)
    - Благодаря CASCADE DELETE удаляются все связанные записи автоматически
    
    Аргументы:
        kp_id: порядковый номер КП для удаления
        db_path: путь к базе данных
    
    Возвращает:
        True если КП был найден и удалён, False если КП не найдено
    """
    init_schema(db_path)
    conn = _connect(db_path)
    try:
        # КРИТИЧНО: Включаем поддержку FOREIGN KEY (по умолчанию выключена в SQLite)
        conn.execute('PRAGMA foreign_keys = ON')
        
        cur = conn.cursor()
        
        # Проверяем, существует ли КП
        cur.execute('SELECT kp_id FROM KP_offers WHERE kp_id = ?', (kp_id,))
        if not cur.fetchone():
            return False
        
        # Удаляем КП (CASCADE автоматически удалит все связанные записи)
        cur.execute('DELETE FROM KP_offers WHERE kp_id = ?', (kp_id,))
        
        conn.commit()
        print(f"[DB] ✅ КП #{kp_id} успешно удалено из базы данных")
        return True
    
    except Exception as e:
        print(f"[DB] ❌ Ошибка при удалении КП #{kp_id}: {e}")
        return False
    
    finally:
        conn.close()


def clear_all_plates_data(db_path: str = DEFAULT_DB) -> Dict[str, int]:
    """
    Полностью очищает ВСЕ данные о плитах из базы данных.
    
    Простыми словами:
    - Удаляет ВСЕ КП (и в работе, и выполненные, и отклонённые)
    - Удаляет ВСЕ плиты (и невыполненные, и выполненные)
    - Удаляет ВСЕ остатки от резки
    - Сбрасывает счётчики AUTOINCREMENT
    - Это как полностью очистить завод от всех заказов и начать с нуля
    
    ⚠️ ВНИМАНИЕ: Это необратимая операция!
    
    Возвращает:
        Словарь с количеством удалённых записей из каждой таблицы
    """
    init_schema(db_path)
    conn = _connect(db_path)
    try:
        # Включаем поддержку FOREIGN KEY
        conn.execute('PRAGMA foreign_keys = ON')
        
        cur = conn.cursor()
        
        # Подсчитываем количество записей перед удалением
        cur.execute('SELECT COUNT(*) FROM KP_offers')
        kp_count = cur.fetchone()[0]
        
        cur.execute('SELECT COUNT(*) FROM kp_plates')
        plates_count = cur.fetchone()[0]
        
        cur.execute('SELECT COUNT(*) FROM completed_plates')
        completed_count = cur.fetchone()[0]
        
        cur.execute('SELECT COUNT(*) FROM plate_rests')
        rests_count = cur.fetchone()[0]
        
        cur.execute('SELECT COUNT(*) FROM kp_files')
        files_count = cur.fetchone()[0]
        
        cur.execute('SELECT COUNT(*) FROM kp_meta')
        meta_count = cur.fetchone()[0]
        
        # Удаляем все записи из всех таблиц
        # Порядок важен: сначала зависимые таблицы, потом основную
        cur.execute('DELETE FROM kp_plates')
        cur.execute('DELETE FROM completed_plates')
        cur.execute('DELETE FROM plate_rests')
        cur.execute('DELETE FROM kp_files')
        cur.execute('DELETE FROM kp_meta')
        cur.execute('DELETE FROM KP_offers')
        
        # Сбрасываем счётчики AUTOINCREMENT
        cur.execute('''
            DELETE FROM sqlite_sequence 
            WHERE name IN ('KP_offers', 'kp_plates', 'completed_plates', 
                          'plate_rests', 'kp_files', 'kp_meta')
        ''')
        
        conn.commit()
        
        result = {
            'kp_offers': kp_count,
            'kp_plates': plates_count,
            'completed_plates': completed_count,
            'plate_rests': rests_count,
            'kp_files': files_count,
            'kp_meta': meta_count,
            'total': kp_count + plates_count + completed_count + rests_count + files_count + meta_count
        }
        
        print(f"[DB] ✅ ПОЛНАЯ ОЧИСТКА ВСЕХ ДАННЫХ О ПЛИТАХ:")
        print(f"  - Удалено КП: {kp_count}")
        print(f"  - Удалено плит в работе: {plates_count}")
        print(f"  - Удалено выполненных плит: {completed_count}")
        print(f"  - Удалено остатков: {rests_count}")
        print(f"  - Удалено файлов: {files_count}")
        print(f"  - Удалено метаданных: {meta_count}")
        print(f"  - ВСЕГО ЗАПИСЕЙ: {result['total']}")
        
        return result
    
    except Exception as e:
        print(f"[DB] ❌ Ошибка при полной очистке БД: {e}")
        import traceback
        traceback.print_exc()
        conn.rollback()
        raise
    
    finally:
        conn.close()


def get_db_stats(db_path: str = DEFAULT_DB) -> Dict[str, int]:
    """
    Получает статистику по базе данных.
    
    Возвращает:
        Словарь с количеством записей в каждой таблице
    """
    init_schema(db_path)
    conn = _connect(db_path)
    try:
        cur = conn.cursor()
        
        cur.execute('SELECT COUNT(*) FROM KP_offers')
        kp_count = cur.fetchone()[0]
        
        cur.execute('SELECT COUNT(*) FROM kp_plates')
        plates_count = cur.fetchone()[0]
        
        cur.execute('SELECT COUNT(*) FROM completed_plates')
        completed_count = cur.fetchone()[0]
        
        cur.execute('SELECT COUNT(*) FROM plate_rests')
        rests_count = cur.fetchone()[0]
        
        cur.execute('SELECT COUNT(*) FROM kp_meta WHERE status = "в работе"')
        in_work_count = cur.fetchone()[0]
        
        cur.execute('SELECT COUNT(*) FROM kp_meta WHERE status = "выполнено"')
        completed_kp_count = cur.fetchone()[0]
        
        return {
            'kp_total': kp_count,
            'kp_in_work': in_work_count,
            'kp_completed': completed_kp_count,
            'plates_in_work': plates_count,
            'plates_completed': completed_count,
            'plate_rests': rests_count
        }
    
    finally:
        conn.close()


def clear_all_kp(db_path: str = DEFAULT_DB) -> Dict[str, int]:
    """
    Полностью очищает все таблицы с КП из базы данных.
    
    Простыми словами:
    - Удаляет ВСЕ КП из базы данных
    - Очищает все таблицы: KP_offers, kp_plates, kp_files, kp_meta
    - Сбрасывает счётчики AUTOINCREMENT, чтобы новые КП начинались с 1
    - Это как стереть все записи из всех таблиц Excel и начать заново
    
    Возвращает:
        Словарь с количеством удалённых записей из каждой таблицы
    """
    init_schema(db_path)
    conn = _connect(db_path)
    try:
        # Включаем поддержку FOREIGN KEY
        conn.execute('PRAGMA foreign_keys = ON')
        
        cur = conn.cursor()
        
        # Подсчитываем количество записей перед удалением
        cur.execute('SELECT COUNT(*) FROM KP_offers')
        kp_count = cur.fetchone()[0]
        
        cur.execute('SELECT COUNT(*) FROM kp_plates')
        plates_count = cur.fetchone()[0]
        
        cur.execute('SELECT COUNT(*) FROM kp_files')
        files_count = cur.fetchone()[0]
        
        cur.execute('SELECT COUNT(*) FROM kp_meta')
        meta_count = cur.fetchone()[0]
        
        # Удаляем все записи из всех таблиц
        # Порядок важен: сначала зависимые таблицы, потом основную
        cur.execute('DELETE FROM kp_plates')
        cur.execute('DELETE FROM kp_files')
        cur.execute('DELETE FROM kp_meta')
        cur.execute('DELETE FROM KP_offers')
        
        # Сбрасываем счётчики AUTOINCREMENT
        # Это нужно, чтобы новые КП начинались с номера 1
        cur.execute('DELETE FROM sqlite_sequence WHERE name IN (?, ?, ?, ?)', 
                   ('KP_offers', 'kp_plates', 'kp_files', 'kp_meta'))
        
        conn.commit()
        
        result = {
            'kp_offers': kp_count,
            'kp_plates': plates_count,
            'kp_files': files_count,
            'kp_meta': meta_count
        }
        
        print(f"[DB] ✅ Полная очистка БД завершена:")
        print(f"  - Удалено КП: {kp_count}")
        print(f"  - Удалено записей плит: {plates_count}")
        print(f"  - Удалено файлов: {files_count}")
        print(f"  - Удалено метаданных: {meta_count}")
        
        return result
    
    except Exception as e:
        print(f"[DB] ❌ Ошибка при полной очистке БД: {e}")
        import traceback
        traceback.print_exc()
        raise
    
    finally:
        conn.close()


# ==================== ФУНКЦИИ ДЛЯ ВЫПОЛНЕННЫХ ПЛИТ ====================

def _normalize_plate_name(name: str) -> str:
    """
    Нормализует имя плиты для совпадений (DEPRECATED — оставлен для обратной
    совместимости со старым кодом). Используйте :func:`core.plate_name.canonical`.

    P2: единый источник истины — ``core/plate_name.py``.
    """
    from core import plate_name as _pn
    return _pn.canonical(name)


class UnmovedPlateInfo(TypedDict):
    kp_id: int
    plate_name: str
    qty: int
    length_m: float
    width_m: float
    load_class: int


def move_plates_to_completed(
    kp_id: int,
    plates_to_complete: List[Dict],
    production_day: int,
    db_path: str = DEFAULT_DB,
    plan_ids: Optional[List[str]] = None,
    allow_cross_kp: bool = False,
    *,
    actor: str | None = None,
    return_unmoved: bool = False,
    _external_conn: Optional[sqlite3.Connection] = None,
) -> int | tuple[int, list[UnmovedPlateInfo]]:
    """
    Переносит плиты из kp_plates в completed_plates.
    
    Простыми словами:
    - Берёт список плит, которые нужно отметить как выполненные
    - Записывает их в таблицу completed_plates
    - Уменьшает количество в kp_plates (или удаляет, если qty = 0)
    
    Аргументы:
        kp_id: номер КП, к которому относятся плиты
        plates_to_complete: список плит для переноса (словари с plate_name, qty и т.д.)
        production_day: номер дня производства
        db_path: путь к базе данных
        _external_conn: если задано — функция работает в существующей транзакции
            переданного соединения, не делает commit/rollback и не закрывает conn.
            Все исключения пробрасываются вызывающему слою, чтобы он мог
            откатить свою транзакцию целиком (P0: атомарность complete_day).
    
    Возвращает:
        Количество перенесённых плит (сумма qty)
    """
    own_conn = _external_conn is None
    if own_conn:
        init_schema(db_path)
        conn = _connect(db_path)
    else:
        conn = _external_conn
    completed_count = 0
    unmoved_plates: list[UnmovedPlateInfo] = []
    # #region agent log
    import json as _json4
    _debug_log4 = _DEBUG_LOG
    # #endregion
    
    try:
        if own_conn:
            conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()
        completed_date = datetime.now().strftime('%d.%m.%Y')
        
        # Целевые подстроки для логов (плиты, которые не списались у пользователя)
        _target_substrings = (
            '61,1-12-10',
            '61,2-12-8п',
            '36,6-6,65',
            '64,8-12-12,5',
            '78,1-12-10',
            '78,1-12-12,5',
            # Дополнительные плиты из текущего кейса пользователя
            '25,4-12-8п',
            '43-12-8п',
            '63,9-12-8п',
            '45-7-6п',
            '60-6,65-8п',
            '59,8-12-8п',
        )
        _is_target = lambda n: any(s in (n or '') for s in _target_substrings)

        def find_one_row(plate_name, length_m, width_m, load_class, prefer_kp_id, length_dm_raw=None):
            """Находит одну строку в kp_plates для списания (id, kp_id, plate_name, width_m, qty)."""
            # #region agent log (плиты 61,8-5 / 45-7 / 37,9-9: вход в поиск)
            _trace_substrings = ('61,8-5', '45-7', '37,9-9')
            if any(sub in (plate_name or '') for sub in _trace_substrings):
                try:
                    _log_path = _DEBUG_LOG_8E9428
                    with open(_log_path, 'a', encoding='utf-8') as _f:
                        _f.write(__import__('json').dumps({"sessionId": "8e9428", "hypothesisId": "H_find_row_in", "location": "kp_db:find_one_row:entry", "message": "find_one_row for 61,8-5/45-7/37,9-9", "data": {"plate_name": (plate_name or "")[:80], "length_dm_raw": length_dm_raw, "prefer_kp_id": prefer_kp_id}, "timestamp": __import__("time").time() * 1000}, ensure_ascii=False) + "\n")
                except Exception:
                    pass
                # #endregion
            # ИСПРАВЛЕНИЕ: Подготовка фильтра по ширине для шагов 4-7
            # Предотвращает списание вторичных резов (с другой шириной)
            # из заказанных плит с совпадающей длиной
            _w_clause = 'AND ABS(width_m - ?) < 0.05' if (width_m and width_m > 0) else ''
            _w_params = (width_m,) if (width_m and width_m > 0) else ()
            # ИСПРАВЛЕНИЕ 2: Ищем и 'в плане', и 'в производстве'
            # При разбиении записи assign_plates_to_plan, часть плит может
            # остаться со status='в производстве', но всё равно быть в треках
            _status_filter = "status IN ('в плане', 'в производстве')"
            # 0) По kp_id + length_dm_raw (точное совпадение марки — 59,81 не путается с 59,84)
            if length_dm_raw and str(length_dm_raw).strip():
                _ldr = str(length_dm_raw).strip()
                cur.execute(f'''
                    SELECT id, kp_id, plate_name, width_m, qty, nomenclature_id FROM kp_plates
                    WHERE kp_id = ? AND length_dm_raw = ? AND qty > 0 AND {_status_filter} {_w_clause}
                    LIMIT 1
                '''.replace('  ', ' '), (prefer_kp_id, _ldr) + _w_params)
                row = cur.fetchone()
                if row:
                    return row
            # 1) По kp_id + plate_name (точное)
            cur.execute(f'''
                SELECT id, kp_id, plate_name, width_m, qty, nomenclature_id FROM kp_plates
                WHERE kp_id = ? AND plate_name = ? AND qty > 0 AND {_status_filter}
                LIMIT 1
            ''', (prefer_kp_id, plate_name))
            row = cur.fetchone()
            if row:
                return row
            # 1.5) P2: каноническое сравнение — снимает префикс «Плиты » и пробелы.
            # Покрывает оба направления: plate_name хранит «Плиты ПБ …», ищем «ПБ …»
            # и наоборот. Делаем в Python, чтобы не плодить sql-функции.
            from core import plate_name as _pn
            canon = _pn.canonical(plate_name)
            if canon:
                cur.execute(f'''
                    SELECT id, kp_id, plate_name, width_m, qty, nomenclature_id FROM kp_plates
                    WHERE kp_id = ? AND qty > 0 AND {_status_filter}
                ''', (prefer_kp_id,))
                for cand in cur.fetchall():
                    if _pn.canonical(cand[2]) == canon:
                        return cand
            # 2) Нормализованное имя (DEPRECATED, теперь покрыто шагом 1.5)
            normalized_name = _normalize_plate_name(plate_name)
            if normalized_name and normalized_name != plate_name:
                cur.execute(f'''
                    SELECT id, kp_id, plate_name, width_m, qty, nomenclature_id FROM kp_plates
                    WHERE kp_id = ? AND plate_name = ? AND qty > 0 AND {_status_filter}
                    LIMIT 1
                ''', (prefer_kp_id, normalized_name))
                row = cur.fetchone()
                if row:
                    return row
            # 2.5) Эквивалентные имена: оптимизатор подставляет 59,9 вместо 59,8, 61,1 вместо 61,2
            _equiv = []
            if plate_name and '59,9-12' in plate_name:
                _equiv.append(plate_name.replace('59,9-12', '59,8-12'))
            if plate_name and '59,8-12' in plate_name:
                _equiv.append(plate_name.replace('59,8-12', '59,9-12'))
            if plate_name and '61,1-12' in plate_name:
                _equiv.append(plate_name.replace('61,1-12', '61,2-12'))
            if plate_name and '61,2-12' in plate_name:
                _equiv.append(plate_name.replace('61,2-12', '61,1-12'))
            for eq in _equiv:
                if eq and eq != plate_name:
                    cur.execute(f'''
                        SELECT id, kp_id, plate_name, width_m, qty, nomenclature_id FROM kp_plates
                        WHERE kp_id = ? AND plate_name = ? AND qty > 0 AND {_status_filter}
                        LIMIT 1
                    ''', (prefer_kp_id, eq))
                    row = cur.fetchone()
                    if row:
                        return row
            # 2.55) 61,1/61,2-12: в предпочитаемом КП не нашли — ищем по длине 6.11±0.02 в любом КП
            # (в БД может быть 61,2-12-8п в другом КП, а в плане ушли 61,1 из-за оптимизатора)
            if length_m and (5.9 <= length_m <= 6.2) and (plate_name or '').strip():
                if '61,1-12' in plate_name or '61,2-12' in plate_name:
                    cur.execute(f'''
                        SELECT id, kp_id, plate_name, width_m, qty, nomenclature_id FROM kp_plates
                        WHERE {_status_filter} AND qty > 0
                          AND ABS(length_m - ?) < 0.02 {_w_clause} AND load_class = ?
                        ORDER BY CASE WHEN kp_id = ? THEN 0 ELSE 1 END, id
                        LIMIT 1
                    ''', (length_m, *_w_params, load_class, prefer_kp_id))
                    row = cur.fetchone()
                    if row:
                        try:
                            with open(_DEBUG_LOG, 'a', encoding='utf-8') as _f:
                                _f.write(__import__('json').dumps({
                                    "hypothesisId": "H_61_cross_kp",
                                    "location": "kp_db:find_one_row:step_2.55",
                                    "message": "61,1/61,2 найдена в другом КП по длине",
                                    "data": {"requested_plate_name": plate_name, "requested_length": length_m, "prefer_kp_id": prefer_kp_id, "found_kp_id": row[1], "found_plate_name": row[2]},
                                    "timestamp": __import__('time').time()
                                }, ensure_ascii=False) + '\n')
                        except Exception:
                            pass
                        return row
            # 2.6) Общий допуск по длине (±0.02м) для любых плит (этап 5 плана)
            if length_m:
                length_tolerance = 0.02
                cur.execute(f'''
                    SELECT id, kp_id, plate_name, width_m, qty, nomenclature_id FROM kp_plates
                    WHERE kp_id = ? AND {_status_filter} AND qty > 0
                      AND ABS(length_m - ?) < ? {_w_clause} AND load_class = ?
                    LIMIT 1
                ''', (prefer_kp_id, length_m, length_tolerance, *_w_params, load_class))
                row = cur.fetchone()
                if row:
                    try:
                        with open(_DEBUG_LOG, 'a', encoding='utf-8') as _f:
                            _f.write(__import__('json').dumps({
                                "hypothesisId": "H_length_tolerance",
                                "location": "kp_db:find_one_row:step_2.6",
                                "message": "найдена плита по допуску длины",
                                "data": {
                                    "requested_length": length_m,
                                    "found_plate_name": row[2],
                                    "found_kp_id": row[1],
                                    "prefer_kp_id": prefer_kp_id
                                },
                                "timestamp": __import__('time').time()
                            }, ensure_ascii=False) + '\n')
                    except Exception:
                        pass
                    return row
            # 3) По размерам в prefer_kp_id
            if length_m and width_m:
                cur.execute(f'''
                    SELECT id, kp_id, plate_name, width_m, qty, nomenclature_id FROM kp_plates
                    WHERE kp_id = ? AND {_status_filter} AND qty > 0
                      AND ABS(length_m - ?) < 0.02 AND ABS(width_m - ?) < 0.01 AND load_class = ?
                    LIMIT 1
                ''', (prefer_kp_id, length_m, width_m, load_class))
                row = cur.fetchone()
                if row:
                    return row
            # 4) По длине + ширине + нагрузке в prefer_kp_id
            if length_m:
                cur.execute(f'''
                    SELECT id, kp_id, plate_name, width_m, qty, nomenclature_id FROM kp_plates
                    WHERE kp_id = ? AND {_status_filter} AND qty > 0
                      AND ABS(length_m - ?) < 0.02 {_w_clause} AND load_class = ?
                    LIMIT 1
                ''', (prefer_kp_id, length_m, *_w_params, load_class))
                row = cur.fetchone()
                if row:
                    return row
            # 5) По plan_ids + длина + ширина + нагрузка (только если разрешено списание с другого КП)
            if allow_cross_kp and length_m and plan_ids:
                placeholders = ','.join('?' * len(plan_ids))
                cur.execute(f'''
                    SELECT id, kp_id, plate_name, width_m, qty, nomenclature_id FROM kp_plates
                    WHERE plan_id IN ({placeholders}) AND {_status_filter} AND qty > 0
                      AND ABS(length_m - ?) < 0.02 {_w_clause} AND load_class = ?
                    ORDER BY CASE WHEN kp_id = ? THEN 0 ELSE 1 END, id
                    LIMIT 1
                ''', (*plan_ids, length_m, *_w_params, load_class, prefer_kp_id))
                row = cur.fetchone()
                if row:
                    return row
            # 6) По длине + ширине + нагрузке в любом КП (только если разрешено)
            if allow_cross_kp and length_m:
                cur.execute(f'''
                    SELECT id, kp_id, plate_name, width_m, qty, nomenclature_id FROM kp_plates
                    WHERE {_status_filter} AND qty > 0
                      AND ABS(length_m - ?) < 0.02 {_w_clause} AND load_class = ?
                    ORDER BY CASE WHEN kp_id = ? THEN 0 ELSE 1 END, id
                    LIMIT 1
                ''', (length_m, *_w_params, load_class, prefer_kp_id))
                row = cur.fetchone()
                if row:
                    return row
            # 7) ФОЛБЭК: по длине + ширине + нагрузке + КП
            if length_m:
                cur.execute(f'''
                    SELECT id, kp_id, plate_name, width_m, qty, nomenclature_id FROM kp_plates
                    WHERE qty > 0 AND {_status_filter}
                      AND ABS(length_m - ?) < 0.02 {_w_clause} AND load_class = ?
                      AND kp_id = ?
                    ORDER BY id
                    LIMIT 1
                ''', (length_m, *_w_params, load_class, prefer_kp_id))
                row = cur.fetchone()
                if row:
                    # #region agent log: сработал фолбэк по длине без статуса
                    try:
                        with open(_debug_log4, 'a', encoding='utf-8') as _f4:
                            _f4.write(_json4.dumps({
                                "hypothesisId": "H10",
                                "location": "kp_db:move_plates_to_completed",
                                "message": "Фолбэк: найдена плита без фильтра по статусу",
                                "data": {
                                    "requested_plate_name": plate_name,
                                    "length_m": length_m,
                                    "width_m": width_m,
                                    "load_class": load_class,
                                    "prefer_kp_id": prefer_kp_id,
                                    "row": {
                                        "id": row[0],
                                        "kp_id": row[1],
                                        "plate_name": row[2],
                                        "width_m": row[3],
                                        "qty": row[4],
                                    },
                                },
                                "timestamp": __import__('time').time(),
                            }, ensure_ascii=False) + '\n')
                    except Exception:
                        pass
                    # #endregion
                    return row
            return None

        for plate in plates_to_complete:
            plate_name = plate.get('plate_name', '')
            qty_remaining = plate.get('qty', 1)
            length_m = plate.get('length_m', 0)
            width_m = plate.get('width_m', 0)
            load_class = plate.get('load_class', 800)
            length_dm_raw = plate.get('length_dm_raw') or ''
            kp_plate_id = plate.get('kp_plate_id')
            
            if not plate_name:
                continue

            # #region agent log: целевые плиты на входе в списание
            if _is_target(plate_name):
                cur.execute("SELECT plate_name, length_m, width_m, load_class, qty, status FROM kp_plates WHERE kp_id = ? AND status IN ('в плане', 'в производстве') AND qty > 0 LIMIT 10", (kp_id,))
                _rows_in_db = [dict(zip(('plate_name', 'length_m', 'width_m', 'load_class', 'qty', 'status'), r)) for r in cur.fetchall()]
                with open(_debug_log4, 'a', encoding='utf-8') as _f4:
                    _f4.write(_json4.dumps({"hypothesisId": "H5", "location": "kp_db:move_plates_to_completed", "message": "Целевая плита на входе", "data": {"kp_id": kp_id, "plate_name": plate_name, "length_m": length_m, "width_m": width_m, "load_class": load_class, "qty": qty_remaining, "rows_in_db_for_kp": _rows_in_db}, "timestamp": __import__('time').time()}, ensure_ascii=False) + '\n')
            # #endregion
            # #region agent log H_59_8_request: запрос на списание 59,8
            if '59,8-12-8п' in (plate_name or ''):
                try:
                    with open(_debug_log4, 'a', encoding='utf-8') as _f59:
                        _f59.write(_json4.dumps({"hypothesisId": "H_59_8_request", "location": "kp_db:move_plates_to_completed", "message": "Запрос списания 59,8-12-8п", "data": {"kp_id": kp_id, "qty_requested": qty_remaining, "length_m": length_m, "width_m": width_m}, "timestamp": __import__('time').time()}, ensure_ascii=False) + '\n')
                except Exception:
                    pass
            # #endregion
            # #region agent log H_61_2_request: запрос на списание 61,2
            if '61,2-12-8п' in (plate_name or '') or '61,1-12-8п' in (plate_name or ''):
                try:
                    with open(_debug_log4, 'a', encoding='utf-8') as _f61:
                        _f61.write(_json4.dumps({"hypothesisId": "H_61_2_request", "location": "kp_db:move_plates_to_completed", "message": "Запрос списания 61,2/61,1-12-8п", "data": {"kp_id": kp_id, "plate_name": plate_name, "qty_requested": qty_remaining, "length_m": length_m, "width_m": width_m}, "timestamp": __import__('time').time()}, ensure_ascii=False) + '\n')
                except Exception:
                    pass
            # #endregion

            # Списываем по одной строке за раз, чтобы добирать остаток из других КП (не оставлять 1 шт «в плане»)
            current_plate_name = plate_name
            current_width_m = width_m
            while qty_remaining > 0:
                row = None
                if kp_plate_id:
                    try:
                        if plan_ids:
                            placeholders = ','.join('?' * len(plan_ids))
                            cur.execute(f'''
                                SELECT id, kp_id, plate_name, width_m, qty, nomenclature_id
                                FROM kp_plates
                                WHERE id = ?
                                  AND plan_id IN ({placeholders})
                                  AND day_number = ?
                                  AND status IN ('в плане', 'в производстве')
                                  AND qty > 0
                                LIMIT 1
                            ''', (int(kp_plate_id), *plan_ids, int(production_day)))
                        else:
                            cur.execute('''
                                SELECT id, kp_id, plate_name, width_m, qty, nomenclature_id
                                FROM kp_plates
                                WHERE id = ?
                                  AND day_number = ?
                                  AND status IN ('в плане', 'в производстве')
                                  AND qty > 0
                                LIMIT 1
                            ''', (int(kp_plate_id), int(production_day)))
                        row = cur.fetchone()
                    except Exception:
                        row = None
                if row is None and kp_plate_id:
                    break
                if row is None:
                    row = find_one_row(current_plate_name, length_m, current_width_m, load_class, kp_id, length_dm_raw=length_dm_raw)
                if not row:
                    # #region agent log (61,8-5 / 45-7 / 37,9-9: строка не найдена)
                    if any(sub in (current_plate_name or '') for sub in ('61,8-5', '45-7', '37,9-9')):
                        try:
                            _log_path = _DEBUG_LOG_8E9428
                            with open(_log_path, 'a', encoding='utf-8') as _f:
                                _f.write(__import__('json').dumps({"sessionId": "8e9428", "hypothesisId": "H_find_row_not_found", "location": "kp_db:move_plates_to_completed", "message": "find_one_row returned None for 61,8-5/45-7/37,9-9", "data": {"plate_name": (current_plate_name or "")[:80], "length_dm_raw": length_dm_raw, "kp_id": kp_id, "qty_remaining": qty_remaining}, "timestamp": __import__("time").time() * 1000}, ensure_ascii=False) + "\n")
                        except Exception:
                            pass
                    # #endregion
                    break
                row_id, row_kp_id, row_plate_name, row_width_m, row_qty, row_nomenclature_id = row
                deduct = min(qty_remaining, row_qty)
                cur.execute('UPDATE kp_plates SET qty = qty - ? WHERE id = ?', (deduct, row_id))
                qty_remaining -= deduct
                completed_count += deduct
                current_plate_name = row_plate_name
                current_width_m = row_width_m
                cur.execute('''
                    INSERT INTO completed_plates (
                        kp_id, plate_name, length_m, width_m, load_class,
                        qty, completed_date, production_day, nomenclature_id
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (row_kp_id, row_plate_name, length_m, row_width_m, load_class, deduct, completed_date, production_day, row_nomenclature_id))
                _audit_append(
                    cur,
                    plate_id=row_id,
                    kp_id=row_kp_id,
                    plate_name=row_plate_name,
                    plan_id=(plan_ids[0] if plan_ids else None),
                    day_number=production_day,
                    from_status='в плане',
                    to_status='completed',
                    qty=deduct,
                    reason='completed',
                    actor=actor,
                )
                # #region agent log (61,8-5 / 45-7 / 37,9-9: списание успешно, какая строка в БД)
                if any(sub in (plate_name or '') for sub in ('61,8-5', '45-7', '37,9-9')):
                    try:
                        _log_path = _DEBUG_LOG_8E9428
                        with open(_log_path, 'a', encoding='utf-8') as _f:
                            _f.write(__import__('json').dumps({"sessionId": "8e9428", "hypothesisId": "H_find_row_found", "location": "kp_db:move_plates_to_completed:deduct", "message": "plate found and deducted for 61,8-5/45-7/37,9-9", "data": {"requested_plate_name": (plate_name or "")[:80], "db_plate_name": (row_plate_name or "")[:80], "nomenclature_id": row_nomenclature_id, "deduct": deduct}, "timestamp": __import__("time").time() * 1000}, ensure_ascii=False) + "\n")
                    except Exception:
                        pass
                # #endregion
                # #region agent log: успешное списание
                try:
                    with open(_debug_log4, 'a', encoding='utf-8') as _f4:
                        _f4.write(_json4.dumps({"hypothesisId": "H_postfix", "location": "kp_db:move_plates_to_completed:success", "message": "Плита списана", "data": {"kp_id": row_kp_id, "plate_name": row_plate_name, "length_m": length_m, "width_m": row_width_m, "load_class": load_class, "deduct": deduct, "requested_width": width_m}, "timestamp": __import__('time').time(), "runId": "post-fix"}, ensure_ascii=False) + '\n')
                except Exception:
                    pass
                # #endregion
                if deduct > 0 and (row_kp_id != kp_id):
                    print(f"[DB] ⚠️ Плита списана из КП #{row_kp_id}: {row_plate_name} (qty={deduct})")

            if qty_remaining > 0:
                unmoved_plates.append(
                    {
                        "kp_id": int(kp_id),
                        "plate_name": str(current_plate_name or plate_name or ""),
                        "qty": int(qty_remaining),
                        "length_m": float(length_m or 0),
                        "width_m": float(current_width_m or width_m or 0),
                        "load_class": int(load_class or 0),
                    }
                )
                print(f"[DB] ⚠️ Не найдена плита для списания: КП #{kp_id}, {current_plate_name} (width={current_width_m}, осталось qty={qty_remaining})")
                # #region agent log: плита не найдена в БД (всегда логируем для отладки)
                cur.execute("SELECT plate_name, length_m, width_m, load_class, qty, status FROM kp_plates WHERE kp_id = ? AND status IN ('в плане', 'в производстве') AND qty > 0 LIMIT 10", (kp_id,))
                _rows_in_db = [dict(zip(('plate_name', 'length_m', 'width_m', 'load_class', 'qty', 'status'), r)) for r in cur.fetchall()]
                with open(_debug_log4, 'a', encoding='utf-8') as _f4:
                    _f4.write(_json4.dumps({"hypothesisId": "H_postfix", "location": "kp_db:move_plates_to_completed", "message": "Плита НЕ НАЙДЕНА в БД (width-check)", "data": {"kp_id": kp_id, "plate_name": current_plate_name, "length_m": length_m, "width_m": current_width_m, "load_class": load_class, "qty": qty_remaining, "rows_in_db_for_kp": _rows_in_db}, "timestamp": __import__('time').time(), "runId": "post-fix"}, ensure_ascii=False) + '\n')
                # #endregion
                # #region agent log H_59_8: для 59,8 — ищем ВСЕ строки в БД (любой kp_id)
                if '59,8-12-8п' in (current_plate_name or ''):
                    cur.execute("SELECT id, kp_id, plate_name, length_m, width_m, qty, status FROM kp_plates WHERE (plate_name LIKE '%59,8-12%' OR plate_name LIKE '%59,9-12%') AND status IN ('в плане', 'в производстве') AND qty > 0")
                    _all_59 = [{"id": r[0], "kp_id": r[1], "plate_name": r[2], "length_m": r[3], "width_m": r[4], "qty": r[5], "status": r[6]} for r in cur.fetchall()]
                    try:
                        with open(_debug_log4, 'a', encoding='utf-8') as _f59:
                            _f59.write(_json4.dumps({"hypothesisId": "H_59_8_remain", "location": "kp_db:move_plates_to_completed:qty_remain", "message": "59,8 не списано: остаток qty, все строки в БД", "data": {"requested_kp_id": kp_id, "qty_remaining": qty_remaining, "all_rows_59x12_in_db": _all_59}, "timestamp": __import__('time').time()}, ensure_ascii=False) + '\n')
                    except Exception:
                        pass
                # #endregion
        
        # Удаляем записи с qty <= 0
        cur.execute('DELETE FROM kp_plates WHERE qty <= 0')

        # #region agent log
        try:
            cur.execute(
                """
                SELECT status, COALESCE(day_number, -1), COUNT(*), COALESCE(SUM(qty), 0)
                FROM kp_plates
                WHERE kp_id = ?
                  AND status IN ('в плане', 'в производстве')
                  AND qty > 0
                GROUP BY status, COALESCE(day_number, -1)
                ORDER BY status, COALESCE(day_number, -1)
                """,
                (kp_id,),
            )
            _remaining_for_kp = [
                {
                    "status": _row[0],
                    "day_number": None if int(_row[1]) == -1 else int(_row[1]),
                    "rows": int(_row[2] or 0),
                    "qty": int(_row[3] or 0),
                }
                for _row in cur.fetchall()
            ]
            with open(_DEBUG_AGENT_LOG, 'a', encoding='utf-8') as _agent_f:
                _agent_f.write(_json4.dumps({
                    "sessionId": "ebb546",
                    "runId": "pre-fix",
                    "hypothesisId": "H5",
                    "location": "core/kp_db.py:move_plates_to_completed:exit",
                    "message": "Итог move_plates_to_completed по КП",
                    "data": {
                        "kp_id": kp_id,
                        "production_day": production_day,
                        "plan_ids": plan_ids,
                        "requested_items": len(plates_to_complete or []),
                        "completed_count": completed_count,
                        "unmoved_qty": sum(int(x.get("qty") or 0) for x in unmoved_plates),
                        "unmoved_sample": unmoved_plates[:10],
                        "remaining_for_kp": _remaining_for_kp,
                    },
                    "timestamp": int(__import__('time').time() * 1000),
                }, ensure_ascii=False) + '\n')
        except Exception:
            pass
        # #endregion
        
        if own_conn:
            conn.commit()
        print(f"[DB] ✅ Перенесено {completed_count} плит в completed_plates (КП #{kp_id}, день {production_day})")
        if return_unmoved:
            return completed_count, unmoved_plates
        return completed_count
        
    except Exception as e:
        if own_conn:
            print(f"[DB] ❌ Ошибка при переносе плит: {e}")
            conn.rollback()
            if return_unmoved:
                return 0, []
            return 0
        # В режиме внешней транзакции пробрасываем — caller сделает rollback своей транзакции.
        raise
    
    finally:
        if own_conn:
            conn.close()


# ==================== ФУНКЦИИ ДЛЯ ОСТАТКОВ ПЛИТ ====================

def create_plate_rest(
    kp_id: int,
    source_plate_name: str,
    rest_width_mm: int,
    length_m: float,
    production_day: int,
    qty: int = 1,
    db_path: str = DEFAULT_DB,
    *,
    _external_conn: Optional[sqlite3.Connection] = None,
) -> int:
    """
    Создает запись об остатке плиты.

    Простыми словами:
    - При продольном резе плиты образуется остаток
    - Эта функция сохраняет информацию об остатке в БД
    - Остаток можно использовать для других заказов

    Аргументы:
        kp_id: номер КП, при выполнении которого образовался остаток
        source_plate_name: имя исходной плиты (из которой вырезали)
        rest_width_mm: ширина остатка в мм
        length_m: длина остатка в метрах
        production_day: номер дня производства
        qty: количество остатков (по умолчанию 1)
        db_path: путь к базе данных
        _external_conn: если задано — функция работает в существующей
            транзакции переданного соединения (P0/P6). Без commit/rollback/close.

    Возвращает:
        ID созданной записи или 0 при ошибке
    """
    own_conn = _external_conn is None
    if own_conn:
        init_schema(db_path)
        conn = _connect(db_path)
    else:
        conn = _external_conn

    try:
        if own_conn:
            conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()
        created_date = datetime.now().strftime('%d.%m.%Y')

        cur.execute('''
            INSERT INTO plate_rests (
                kp_id, source_plate_name, rest_width_mm, length_m,
                qty, status, created_date, production_day
            ) VALUES (?, ?, ?, ?, ?, 'available', ?, ?)
        ''', (
            kp_id,
            source_plate_name,
            rest_width_mm,
            length_m,
            qty,
            created_date,
            production_day,
        ))

        rest_id = cur.lastrowid
        if own_conn:
            conn.commit()
        print(f"[DB] ✅ Создан остаток #{rest_id}: {rest_width_mm}мм x {length_m}м (КП #{kp_id})")
        return rest_id

    except Exception as e:
        if own_conn:
            print(f"[DB] ❌ Ошибка при создании остатка: {e}")
            conn.rollback()
            return 0
        # Внешняя транзакция: пробрасываем для отката caller'ом
        raise

    finally:
        if own_conn:
            conn.close()


def get_available_rests(
    kp_id: int = None,
    db_path: str = DEFAULT_DB
) -> List[Dict]:
    """
    Возвращает список доступных остатков.
    
    Простыми словами:
    - Получает все остатки со статусом 'available'
    - Можно фильтровать по номеру КП
    - Возвращает список словарей с информацией об остатках
    
    Аргументы:
        kp_id: номер КП для фильтрации (None = все КП)
        db_path: путь к базе данных
    
    Возвращает:
        Список словарей с информацией об остатках
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        
        if kp_id:
            cur.execute('''
                SELECT 
                    pr.id, pr.kp_id, pr.source_plate_name, pr.rest_width_mm,
                    pr.length_m, pr.qty, pr.status, pr.created_date,
                    pr.production_day, ko.customer_name
                FROM plate_rests pr
                LEFT JOIN KP_offers ko ON pr.kp_id = ko.kp_id
                WHERE pr.status = 'available' AND pr.kp_id = ?
                ORDER BY pr.created_date DESC, pr.rest_width_mm DESC
            ''', (kp_id,))
        else:
            cur.execute('''
                SELECT 
                    pr.id, pr.kp_id, pr.source_plate_name, pr.rest_width_mm,
                    pr.length_m, pr.qty, pr.status, pr.created_date,
                    pr.production_day, ko.customer_name
                FROM plate_rests pr
                LEFT JOIN KP_offers ko ON pr.kp_id = ko.kp_id
                WHERE pr.status = 'available'
                ORDER BY pr.created_date DESC, pr.rest_width_mm DESC
            ''')
        
        return [dict(row) for row in cur.fetchall()]
        
    finally:
        conn.close()


def mark_rest_as_used(
    rest_id: int,
    db_path: str = DEFAULT_DB
) -> bool:
    """
    Помечает остаток как использованный.
    
    Простыми словами:
    - Когда остаток используется во вторичном резе
    - Его статус меняется на 'used'
    - Он больше не будет показываться как доступный
    
    Аргументы:
        rest_id: ID остатка в БД
        db_path: путь к базе данных
    
    Возвращает:
        True если успешно, False при ошибке
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()
        used_date = datetime.now().strftime('%d.%m.%Y')
        
        cur.execute('''
            UPDATE plate_rests
            SET status = 'used', used_date = ?
            WHERE id = ? AND status = 'available'
        ''', (used_date, rest_id))
        
        if cur.rowcount > 0:
            conn.commit()
            print(f"[DB] ✅ Остаток #{rest_id} помечен как использованный")
            return True
        else:
            print(f"[DB] ⚠️ Остаток #{rest_id} не найден или уже использован")
            return False
        
    except Exception as e:
        print(f"[DB] ❌ Ошибка при обновлении остатка: {e}")
        conn.rollback()
        return False
    
    finally:
        conn.close()


def complete_plate_rest(
    rest_id: int,
    production_day: int,
    db_path: str = DEFAULT_DB
) -> bool:
    """
    Помечает остаток как выполненный (произведённый).
    
    Простыми словами:
    - Когда остаток изготавливается как отдельная плита
    - Его статус меняется на 'completed'
    - Записывается день производства
    
    Аргументы:
        rest_id: ID остатка в БД
        production_day: номер дня производства
        db_path: путь к базе данных
    
    Возвращает:
        True если успешно, False при ошибке
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()
        used_date = datetime.now().strftime('%d.%m.%Y')
        
        cur.execute('''
            UPDATE plate_rests
            SET status = 'completed', used_date = ?, production_day = ?
            WHERE id = ? AND status = 'available'
        ''', (used_date, production_day, rest_id))
        
        if cur.rowcount > 0:
            conn.commit()
            print(f"[DB] ✅ Остаток #{rest_id} выполнен (день {production_day})")
            return True
        else:
            print(f"[DB] ⚠️ Остаток #{rest_id} не найден или уже обработан")
            return False
        
    except Exception as e:
        print(f"[DB] ❌ Ошибка при завершении остатка: {e}")
        conn.rollback()
        return False
    
    finally:
        conn.close()


def discard_plate_rest(
    rest_id: int,
    db_path: str = DEFAULT_DB
) -> bool:
    """
    Списывает остаток (брак или утилизация).
    
    Простыми словами:
    - Когда остаток больше не нужен (брак, повреждение)
    - Его статус меняется на 'discarded'
    - Остаток удаляется из доступных
    
    Аргументы:
        rest_id: ID остатка в БД
        db_path: путь к базе данных
    
    Возвращает:
        True если успешно, False при ошибке
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()
        used_date = datetime.now().strftime('%d.%m.%Y')
        
        cur.execute('''
            UPDATE plate_rests
            SET status = 'discarded', used_date = ?
            WHERE id = ? AND status = 'available'
        ''', (used_date, rest_id))
        
        if cur.rowcount > 0:
            conn.commit()
            print(f"[DB] ✅ Остаток #{rest_id} списан")
            return True
        else:
            print(f"[DB] ⚠️ Остаток #{rest_id} не найден или уже обработан")
            return False
        
    except Exception as e:
        print(f"[DB] ❌ Ошибка при списании остатка: {e}")
        conn.rollback()
        return False
    
    finally:
        conn.close()


def get_all_plate_rests(db_path: str = DEFAULT_DB) -> List[Dict]:
    """
    Получает все остатки (для статистики и отчётов).
    
    Возвращает:
        Список словарей с информацией обо всех остатках
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        
        cur.execute('''
            SELECT 
                pr.id, pr.kp_id, pr.source_plate_name, pr.rest_width_mm,
                pr.length_m, pr.qty, pr.status, pr.created_date,
                pr.used_date, pr.production_day, ko.customer_name
            FROM plate_rests pr
            LEFT JOIN KP_offers ko ON pr.kp_id = ko.kp_id
            ORDER BY pr.created_date DESC, pr.status, pr.rest_width_mm DESC
        ''')
        
        return [dict(row) for row in cur.fetchall()]
        
    finally:
        conn.close()


def find_matching_rests(
    length_m: float,
    width_mm: int,
    qty_needed: int,
    db_path: str = DEFAULT_DB
) -> List[Dict]:
    """
    Ищет остатки, из которых можно получить плиту нужного размера.
    
    Простыми словами:
    - Ищет остатки со статусом 'available'
    - Условие: длина остатка >= длина плиты И ширина остатка >= ширина плиты
    - Возвращает список подходящих остатков с информацией о типе совпадения
    
    Аргументы:
        length_m: требуемая длина плиты в метрах
        width_mm: требуемая ширина плиты в мм
        qty_needed: сколько плит нужно
        db_path: путь к базе данных
    
    Возвращает:
        Список словарей с информацией о подходящих остатках:
        - rest_id: ID остатка
        - rest_length: длина остатка
        - rest_width_mm: ширина остатка
        - match_type: 'exact' / 'width_cut' / 'length_cut' / 'both_cuts'
        - cut_cost: себестоимость резов (0 при exact)
        - source_plate_name: название исходной плиты
        - source_kp_id: КП, из которого остаток
    """
    # Константы стоимости резов (из config_and_data)
    LONG_CUT_PRICE_PER_M = 460.0   # Продольный рез, руб/пог.м
    TRANSVERSE_CUT_PRICE = 1200.0  # Поперечный рез, руб/шт
    
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        
        # Ищем остатки, которые >= по обоим размерам
        cur.execute('''
            SELECT 
                pr.id, pr.kp_id, pr.source_plate_name, pr.rest_width_mm,
                pr.length_m, pr.qty, pr.status, pr.created_date,
                pr.production_day, ko.customer_name
            FROM plate_rests pr
            LEFT JOIN KP_offers ko ON pr.kp_id = ko.kp_id
            WHERE pr.status = 'available'
              AND pr.length_m >= ?
              AND pr.rest_width_mm >= ?
            ORDER BY 
                -- Приоритет: сначала точные совпадения, потом минимальные отходы
                CASE 
                    WHEN pr.length_m = ? AND pr.rest_width_mm = ? THEN 0  -- exact
                    WHEN pr.length_m = ? THEN 1  -- только рез по ширине
                    WHEN pr.rest_width_mm = ? THEN 2  -- только рез по длине
                    ELSE 3  -- оба реза
                END,
                (pr.length_m - ?) + (pr.rest_width_mm - ?) / 1000.0  -- минимум отходов
        ''', (length_m, width_mm, length_m, width_mm, length_m, width_mm, length_m, width_mm))
        
        results = []
        qty_collected = 0
        
        for row in cur.fetchall():
            if qty_collected >= qty_needed:
                break
            
            rest = dict(row)
            rest_length = rest['length_m']
            rest_width = rest['rest_width_mm']
            rest_qty = rest['qty']
            
            # Определяем тип совпадения и стоимость резов
            length_match = abs(rest_length - length_m) < 0.01  # точное совпадение по длине
            width_match = rest_width == width_mm  # точное совпадение по ширине
            
            if length_match and width_match:
                match_type = 'exact'
                cut_cost = 0.0
            elif length_match and not width_match:
                match_type = 'width_cut'
                cut_cost = LONG_CUT_PRICE_PER_M * length_m
            elif not length_match and width_match:
                match_type = 'length_cut'
                cut_cost = TRANSVERSE_CUT_PRICE
            else:
                match_type = 'both_cuts'
                cut_cost = LONG_CUT_PRICE_PER_M * length_m + TRANSVERSE_CUT_PRICE
            
            # Сколько можем взять из этого остатка
            can_take = min(rest_qty, qty_needed - qty_collected)
            
            results.append({
                'rest_id': rest['id'],
                'rest_length': rest_length,
                'rest_width_mm': rest_width,
                'rest_qty_available': rest_qty,
                'qty_to_use': can_take,
                'match_type': match_type,
                'cut_cost': cut_cost,
                'source_plate_name': rest['source_plate_name'],
                'source_kp_id': rest['kp_id'],
                'source_customer': rest.get('customer_name', 'неизвестно')
            })
            
            qty_collected += can_take
        
        return results
        
    finally:
        conn.close()


def check_and_update_kp_completion(
    kp_id: int,
    db_path: str = DEFAULT_DB,
    *,
    _external_conn: Optional[sqlite3.Connection] = None,
) -> bool:
    """
    Проверяет, все ли плиты КП выполнены.
    Если да — меняет статус КП на "выполнено".

    Простыми словами:
    - Смотрит, остались ли ещё плиты в kp_plates для данного КП
    - Если плит не осталось (все выполнены) — ставит статус "выполнено"

    Аргументы:
        kp_id: номер КП для проверки
        db_path: путь к базе данных
        _external_conn: если задано — функция работает в существующей транзакции
            переданного соединения (P0). Не делает commit/rollback и не закрывает conn.

    Возвращает:
        True если КП полностью выполнен, False если ещё есть плиты
    """
    own_conn = _external_conn is None
    if own_conn:
        init_schema(db_path)
        conn = _connect(db_path)
    else:
        conn = _external_conn

    try:
        if own_conn:
            conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()

        # Считаем оставшиеся плиты в КП
        cur.execute('SELECT SUM(qty) FROM kp_plates WHERE kp_id = ?', (kp_id,))
        result = cur.fetchone()
        remaining = result[0] if result[0] else 0

        if remaining == 0:
            # Все плиты выполнены — обновляем статус
            cur.execute('''
                UPDATE kp_meta SET status = 'выполнено' WHERE kp_id = ?
            ''', (kp_id,))
            if own_conn:
                conn.commit()
            print(f"[DB] 🎉 КП #{kp_id} полностью выполнен! Статус обновлён.")
            return True

        return False

    finally:
        if own_conn:
            conn.close()


def get_remaining_plates_for_kp(kp_id: int, db_path: str = DEFAULT_DB) -> List[Dict]:
    """
    Получает список оставшихся (невыполненных) плит для КП.
    
    Простыми словами:
    - Возвращает все плиты, которые ещё не выполнены для данного КП
    
    Аргументы:
        kp_id: номер КП
        db_path: путь к базе данных
    
    Возвращает:
        Список словарей с информацией о плитах
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        
        cur.execute('''
            SELECT * FROM kp_plates 
            WHERE kp_id = ? AND qty > 0
            ORDER BY position_number
        ''', (kp_id,))
        
        return [dict(row) for row in cur.fetchall()]
        
    finally:
        conn.close()


# =============================================================================
# ФУНКЦИИ ДЛЯ УПРАВЛЕНИЯ СТАТУСАМИ ПЛИТ
# =============================================================================

def mark_plates_as_planned(
    kp_id: int,
    plate_name: str,
    qty_to_plan: int,
    plan_id: str,
    db_path: str = DEFAULT_DB,
    *,
    actor: str | None = None,
    day_number: Optional[int] = None,
) -> Dict[str, object]:
    """
    Помечает плиты как 'в плане' при сохранении плана.

    ИСПРАВЛЕНО: Теперь обрабатывает ВСЕ записи с одинаковым plate_name,
    а не только первую (убран LIMIT 1).

    Простыми словами:
    - Находит ВСЕ плиты по kp_id и plate_name со статусом 'в производстве'
    - Обрабатывает их по очереди, пока не наберется нужное qty_to_plan
    - Если qty_to_plan = всему qty — просто меняет статус на 'в плане'
    - Если qty_to_plan < qty — разбивает запись на две:
      * Одна с qty_to_plan и статусом 'в плане'
      * Вторая с остатком и статусом 'в производстве'

    Аргументы:
        kp_id: номер КП
        plate_name: название плиты (например, "ПБ 60-12-8п")
        qty_to_plan: сколько плит добавить в план
        plan_id: ID плана (для связи и отката)
        db_path: путь к базе данных
        day_number: номер дня производства (P5). Если задан — пишется в
            ``kp_plates.day_number`` и в audit-лог. ``None`` оставляет
            старое поведение (день не зафиксирован — legacy).

    Возвращает:
        Dict с подробным результатом пометки. ``updated_ids`` — id строк
        kp_plates, которые получили статус «в плане» (нужно вызывающему
        слою, чтобы записать ``kp_plate_id`` в plan.json/items).
    """
    init_schema(db_path)
    conn = _connect(db_path)
    requested_qty = max(int(qty_to_plan or 0), 0)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()
        
        # ИСПРАВЛЕНО: Находим ВСЕ плиты со статусом 'в производстве' (без LIMIT 1)
        cur.execute('''
            SELECT id, qty, position_number, length_m, width_m, load_class,
                   unit_weight, total_weight, discounted_price
            FROM kp_plates
            WHERE kp_id = ? AND plate_name = ? AND status = 'в производстве' AND qty > 0
            ORDER BY id
        ''', (kp_id, plate_name))
        
        rows = cur.fetchall()
        # #region agent log
        if any(_k in (plate_name or "") for _k in ("59,8-12-8п", "50,8-5,3-8п", "50,8-3,2-8п", "63,9-12-8п")):
            _debug_session_write(
                "run1",
                "H4",
                "kp_db:mark_plates_as_planned:rows_loaded",
                "Rows selected for marking as planned",
                {
                    "kp_id": kp_id,
                    "plate_name": plate_name,
                    "qty_to_plan_requested": qty_to_plan,
                    "rows_count": len(rows),
                    "rows": [{"id": r[0], "qty": r[1], "length_m": r[3], "width_m": r[4], "load_class": r[5]} for r in rows],
                },
            )
        # #endregion
        if not rows:
            print(f"[DB] ⚠️ Плита не найдена: КП #{kp_id}, {plate_name} (статус 'в производстве')")
            return {
                'success': False,
                'requested_qty': requested_qty,
                'available_qty': 0,
                'processed_count': 0,
                'remaining_unplanned': requested_qty,
                'split_count': 0,
                'updated_ids': [],
                'id_qty_pairs': [],
                'error': "plate_not_found",
            }
        
        # Подсчитываем общее доступное количество
        total_available = sum(row[1] for row in rows)
        
        if qty_to_plan > total_available:
            print(f"[DB] ⚠️ Запрошено {qty_to_plan}, но доступно только {total_available}")
            qty_to_plan = total_available
        
        # Обрабатываем записи по очереди
        remaining_to_plan = qty_to_plan
        processed_count = 0
        split_count = 0
        updated_ids: List[int] = []
        # P5: id_qty_pairs нужен caller'у, чтобы при записи kp_plate_id в plan.json
        # знать, сколько items приходится на каждый plate_id.
        id_qty_pairs: List[Tuple[int, int]] = []

        for row in rows:
            if remaining_to_plan <= 0:
                break
            
            plate_id, current_qty, pos_num, length_m, width_m, load_class, unit_w, total_w, price = row
            
            if current_qty <= remaining_to_plan:
                # Вся запись идет в план
                cur.execute('''
                    UPDATE kp_plates
                    SET status = 'в плане', plan_id = ?, day_number = ?
                    WHERE id = ?
                ''', (plan_id, day_number, plate_id))
                _audit_append(
                    cur,
                    plate_id=plate_id,
                    kp_id=kp_id,
                    plate_name=plate_name,
                    plan_id=plan_id,
                    day_number=day_number,
                    from_status='в производстве',
                    to_status='в плане',
                    qty=current_qty,
                    reason='planned',
                    actor=actor,
                )
                print(f"[DB] ✅ Плита {plate_name} x{current_qty} помечена как 'в плане' (запись #{plate_id})")
                remaining_to_plan -= current_qty
                processed_count += current_qty
                updated_ids.append(plate_id)
                id_qty_pairs.append((plate_id, current_qty))
            else:
                # Частичная обработка: разбиваем запись
                qty_for_plan = remaining_to_plan
                remaining_in_production = current_qty - qty_for_plan

                # 1. Обновляем существующую запись: уменьшаем qty, помечаем 'в плане'
                cur.execute('''
                    UPDATE kp_plates
                    SET qty = ?, status = 'в плане', plan_id = ?, day_number = ?
                    WHERE id = ?
                ''', (qty_for_plan, plan_id, day_number, plate_id))

                # 2. Создаём новую запись с остатком (статус 'в производстве')
                cur.execute('''
                    INSERT INTO kp_plates (
                        kp_id, position_number, plate_name, length_m, width_m, load_class,
                        qty, unit_weight, total_weight, discounted_price, status, plan_id, day_number
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'в производстве', NULL, NULL)
                ''', (kp_id, pos_num, plate_name, length_m, width_m, load_class,
                      remaining_in_production, unit_w, total_w, price))
                _audit_append(
                    cur,
                    plate_id=plate_id,
                    kp_id=kp_id,
                    plate_name=plate_name,
                    plan_id=plan_id,
                    day_number=day_number,
                    from_status='в производстве',
                    to_status='в плане',
                    qty=qty_for_plan,
                    reason='planned',
                    actor=actor,
                )
                
                print(f"[DB] ✅ Плита {plate_name} разбита: {qty_for_plan} в план, {remaining_in_production} осталось (запись #{plate_id})")
                remaining_to_plan = 0
                processed_count += qty_for_plan
                split_count += 1
                updated_ids.append(plate_id)
                id_qty_pairs.append((plate_id, qty_for_plan))
        
        print(f"[DB] ✅ Итого помечено {processed_count} плит '{plate_name}' как 'в плане' (план {plan_id})")
        # #region agent log
        if any(_k in (plate_name or "") for _k in ("59,8-12-8п", "50,8-5,3-8п", "50,8-3,2-8п", "63,9-12-8п")):
            _debug_session_write(
                "run1",
                "H4",
                "kp_db:mark_plates_as_planned:done",
                "Mark as planned result for target plate",
                {
                    "kp_id": kp_id,
                    "plate_name": plate_name,
                    "processed_count": processed_count,
                    "remaining_to_plan": remaining_to_plan,
                    "plan_id": plan_id,
                },
            )
        # #endregion
        conn.commit()
        return {
            'success': True,
            'requested_qty': requested_qty,
            'available_qty': total_available,
            'processed_count': processed_count,
            'remaining_unplanned': max(requested_qty - processed_count, 0),
            'split_count': split_count,
            'updated_ids': updated_ids,
            'id_qty_pairs': id_qty_pairs,
            'error': None,
        }
        
    except Exception as e:
        print(f"[DB] ❌ Ошибка при пометке плиты как 'в плане': {e}")
        conn.rollback()
        return {
            'success': False,
            'requested_qty': requested_qty,
            'available_qty': 0,
            'processed_count': 0,
            'remaining_unplanned': requested_qty,
            'split_count': 0,
            'updated_ids': [],
            'id_qty_pairs': [],
            'error': str(e),
        }
    
    finally:
        conn.close()


def return_plates_to_production(
    kp_id: int,
    plate_name: str,
    qty: int,
    db_path: str = DEFAULT_DB,
    *,
    actor: str | None = None,
    reason: str = "rejected",
    _external_conn: Optional[sqlite3.Connection] = None,
) -> bool:
    """
    Возвращает плиты обратно в статус 'в производстве'.

    ИСПРАВЛЕНО: Теперь обрабатывает ВСЕ записи с одинаковым plate_name,
    а не только первую (убран LIMIT 1).

    Простыми словами:
    - Используется при браке: бракованные плиты возвращаются в производство
    - Находит ВСЕ плиты со статусом 'в плане' и меняет статус на 'в производстве'
    - Очищает plan_id
    - Обрабатывает нужное количество (qty) по записям

    Аргументы:
        kp_id: номер КП
        plate_name: название плиты
        qty: количество плит для возврата
        db_path: путь к базе данных
        _external_conn: если задано — функция работает в существующей транзакции
            переданного соединения (P0). Не делает commit/rollback и не закрывает conn.
            Все исключения пробрасываются вызывающему слою.

    Возвращает:
        True если успешно, False при ошибке
    """
    own_conn = _external_conn is None
    if own_conn:
        init_schema(db_path)
        conn = _connect(db_path)
    else:
        conn = _external_conn

    try:
        if own_conn:
            conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()

        # ИСПРАВЛЕНО: Находим ВСЕ плиты со статусом 'в плане' (без LIMIT 1)
        cur.execute('''
            SELECT id, qty
            FROM kp_plates
            WHERE kp_id = ? AND plate_name = ? AND status = 'в плане' AND qty > 0
            ORDER BY id
        ''', (kp_id, plate_name))

        rows = cur.fetchall()
        if not rows:
            # Плита могла уже быть возвращена или не была в плане
            print(f"[DB] ⚠️ Плита для возврата не найдена: КП #{kp_id}, {plate_name}")
            return False
        
        # Обрабатываем записи по очереди
        remaining_to_return = qty
        processed_count = 0
        
        for row in rows:
            if remaining_to_return <= 0:
                break
            
            plate_id, current_qty = row
            
            if current_qty <= remaining_to_return:
                # Вся запись возвращается в производство
                cur.execute('''
                    UPDATE kp_plates
                    SET status = 'в производстве', plan_id = NULL
                    WHERE id = ?
                ''', (plate_id,))
                _audit_append(
                    cur,
                    plate_id=plate_id,
                    kp_id=kp_id,
                    plate_name=plate_name,
                    plan_id=None,
                    day_number=None,
                    from_status='в плане',
                    to_status='в производстве',
                    qty=current_qty,
                    reason=reason,
                    actor=actor,
                )
                print(f"[DB] ✅ Плита {plate_name} x{current_qty} возвращена в производство (запись #{plate_id})")
                remaining_to_return -= current_qty
                processed_count += current_qty
            else:
                # Частичный возврат: уменьшаем qty в плане, создаём новую запись «в производстве»
                new_qty_in_plan = current_qty - remaining_to_return
                cur.execute('''
                    UPDATE kp_plates
                    SET qty = ?
                    WHERE id = ?
                ''', (new_qty_in_plan, plate_id))
                cur.execute('''
                    INSERT INTO kp_plates (
                        kp_id, position_number, plate_name, length_m, width_m,
                        load_class, qty, unit_weight, total_weight, discounted_price,
                        status, plan_id
                    )
                    SELECT
                        kp_id, position_number, plate_name, length_m, width_m,
                        load_class, ?, unit_weight, total_weight, discounted_price,
                        'в производстве', NULL
                    FROM kp_plates WHERE id = ?
                ''', (remaining_to_return, plate_id))
                _audit_append(
                    cur,
                    plate_id=plate_id,
                    kp_id=kp_id,
                    plate_name=plate_name,
                    plan_id=None,
                    day_number=None,
                    from_status='в плане',
                    to_status='в производстве',
                    qty=remaining_to_return,
                    reason=reason,
                    actor=actor,
                )
                print(f"[DB] ✅ Частичный возврат: {plate_name} x{remaining_to_return} в производство (осталось в плане: {new_qty_in_plan}, запись #{plate_id})")
                processed_count += remaining_to_return
                remaining_to_return = 0
                break
        
        if own_conn:
            conn.commit()
        print(f"[DB] ✅ Итого возвращено {processed_count} плит '{plate_name}' в производство (КП #{kp_id})")
        return True

    except Exception as e:
        if own_conn:
            print(f"[DB] ❌ Ошибка при возврате плиты в производство: {e}")
            conn.rollback()
            return False
        # Внешняя транзакция: пробрасываем исключение для отката caller'ом
        raise
    
    finally:
        if own_conn:
            conn.close()


def return_plan_plates_to_production(plan_id: str, db_path: str = DEFAULT_DB) -> int:
    """
    Возвращает ВСЕ плиты плана обратно в производство.
    
    Простыми словами:
    - Используется при удалении плана
    - Находит все плиты с данным plan_id
    - Меняет их статус на 'в производстве'
    - Очищает plan_id
    
    Аргументы:
        plan_id: ID плана
        db_path: путь к базе данных
    
    Возвращает:
        Количество возвращённых плит (записей)
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()
        
        # Считаем сколько плит вернём
        cur.execute('''
            SELECT COUNT(*), COALESCE(SUM(qty), 0)
            FROM kp_plates
            WHERE plan_id = ?
        ''', (plan_id,))
        
        count_result = cur.fetchone()
        records_count = count_result[0] if count_result else 0
        qty_total = count_result[1] if count_result else 0
        
        if records_count == 0:
            print(f"[DB] ℹ️ Нет плит для возврата из плана {plan_id}")
            return 0
        
        # Возвращаем все плиты плана в производство
        cur.execute('''
            UPDATE kp_plates
            SET status = 'в производстве', plan_id = NULL
            WHERE plan_id = ?
        ''', (plan_id,))
        
        conn.commit()
        print(f"[DB] ✅ Возвращено {records_count} записей ({qty_total} плит) из плана {plan_id} в производство")
        return records_count
        
    except Exception as e:
        print(f"[DB] ❌ Ошибка при возврате плит плана в производство: {e}")
        conn.rollback()
        return 0
    
    finally:
        conn.close()


def recover_stuck_plates(db_path: str = DEFAULT_DB) -> int:
    """
    Возвращает "застрявшие" плиты обратно в производство.
    
    Простыми словами:
    - Застрявшая плита = статус 'в плане', но не была списана
    - Эти плиты не попали в tracks при планировании
    - Возвращаем их в 'в производстве', чтобы можно было заново запланировать
    
    Когда использовать:
    - После выполнения планов, если часть плит "потерялась"
    - Через команду бота /recover_plates
    
    Аргументы:
        db_path: путь к базе данных
    
    Возвращает:
        Количество восстановленных записей плит
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()
        
        # Сначала посмотрим, сколько плит застряло
        cur.execute('''
            SELECT COUNT(*), COALESCE(SUM(qty), 0)
            FROM kp_plates
            WHERE status = 'в плане'
        ''')
        
        count_result = cur.fetchone()
        records_count = count_result[0] if count_result else 0
        qty_total = count_result[1] if count_result else 0
        
        if records_count == 0:
            print("[DB] ℹ️ Нет застрявших плит для восстановления")
            return 0
        
        # Возвращаем все плиты со статусом 'в плане' в производство
        cur.execute('''
            UPDATE kp_plates
            SET status = 'в производстве', plan_id = NULL
            WHERE status = 'в плане'
        ''')
        
        conn.commit()
        print(f"[DB] ✅ Восстановлено {records_count} записей ({qty_total} плит) из статуса 'в плане' в 'в производстве'")
        return records_count
        
    except Exception as e:
        print(f"[DB] ❌ Ошибка при восстановлении застрявших плит: {e}")
        conn.rollback()
        return 0
    
    finally:
        conn.close()


def return_lost_plates_to_production(
    plan_id: str,
    lost_plates: list,
    db_path: str = DEFAULT_DB
) -> int:
    """
    Возвращает конкретные "потерянные" плиты обратно в производство.
    
    ИСПРАВЛЕНО: Теперь обрабатывает ВСЕ записи с одинаковым plate_name,
    а не только первую (убран LIMIT 1).
    
    Простыми словами:
    - Принимает список плит, которые не попали в tracks
    - Возвращает их статус на 'в производстве'
    - Используется как защита при сохранении плана
    - Обрабатывает нужное количество (qty_lost) по нескольким записям
    
    Аргументы:
        plan_id: ID плана (для логирования)
        lost_plates: список словарей [{'kp_id': X, 'plate_name': Y, 'qty_lost': Z}, ...]
        db_path: путь к базе данных
    
    Возвращает:
        Количество восстановленных записей
    """
    if not lost_plates:
        return 0
    
    init_schema(db_path)
    conn = _connect(db_path)
    total_returned = 0
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        cur = conn.cursor()
        
        for lost in lost_plates:
            kp_id = lost.get('kp_id')
            plate_name = lost.get('plate_name')
            qty_lost = lost.get('qty_lost', 0)
            
            if not kp_id or not plate_name or qty_lost <= 0:
                continue
            
            # ИСПРАВЛЕНО: Ищем ВСЕ плиты со статусом 'в плане' и данным plan_id (без LIMIT 1)
            cur.execute('''
                SELECT id, qty
                FROM kp_plates
                WHERE kp_id = ? AND plate_name = ? AND status = 'в плане' AND plan_id = ?
                ORDER BY id
            ''', (kp_id, plate_name, plan_id))
            
            rows = cur.fetchall()
            if not rows:
                print(f"[DB] ⚠️ Потерянная плита не найдена: КП #{kp_id}, {plate_name}")
                continue
            
            # Обрабатываем записи по очереди
            remaining_to_return = qty_lost
            
            for row in rows:
                if remaining_to_return <= 0:
                    break
                
                plate_id, current_qty = row
                
                if remaining_to_return >= current_qty:
                    # Возвращаем всю запись в производство
                    cur.execute('''
                        UPDATE kp_plates
                        SET status = 'в производстве', plan_id = NULL
                        WHERE id = ?
                    ''', (plate_id,))
                    print(f"[DB] ✅ Возвращена вся запись: {plate_name} x{current_qty} (запись #{plate_id})")
                    remaining_to_return -= current_qty
                    total_returned += 1
                else:
                    # Частичный возврат: уменьшаем qty в плане, создаём новую запись в производстве
                    new_qty_in_plan = current_qty - remaining_to_return
                    
                    # Обновляем существующую запись
                    cur.execute('''
                        UPDATE kp_plates
                        SET qty = ?
                        WHERE id = ?
                    ''', (new_qty_in_plan, plate_id))
                    
                    # Создаём новую запись для возвращённых плит
                    cur.execute('''
                        INSERT INTO kp_plates (
                            kp_id, position_number, plate_name, length_m, width_m,
                            load_class, qty, unit_weight, total_weight, discounted_price,
                            status, plan_id
                        )
                        SELECT 
                            kp_id, position_number, plate_name, length_m, width_m,
                            load_class, ?, unit_weight, total_weight, discounted_price,
                            'в производстве', NULL
                        FROM kp_plates WHERE id = ?
                    ''', (remaining_to_return, plate_id))
                    
                    print(f"[DB] ✅ Частичный возврат: {plate_name} x{remaining_to_return} (осталось в плане: {new_qty_in_plan}, запись #{plate_id})")
                    remaining_to_return = 0
                    total_returned += 1
        
        conn.commit()
        print(f"[DB] ✅ Всего возвращено {total_returned} записей потерянных плит")
        return total_returned
        
    except Exception as e:
        print(f"[DB] ❌ Ошибка при возврате потерянных плит: {e}")
        conn.rollback()
        return 0
    
    finally:
        conn.close()


def get_completed_plates_for_kp(kp_id: int, db_path: str = DEFAULT_DB) -> List[Dict]:
    """
    Получает список выполненных плит для КП.
    
    Простыми словами:
    - Возвращает все плиты, которые уже выполнены для данного КП
    
    Аргументы:
        kp_id: номер КП
        db_path: путь к базе данных
    
    Возвращает:
        Список словарей с информацией о выполненных плитах
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        
        cur.execute('''
            SELECT * FROM completed_plates 
            WHERE kp_id = ?
            ORDER BY completed_date, production_day
        ''', (kp_id,))
        
        return [dict(row) for row in cur.fetchall()]
        
    finally:
        conn.close()


def get_completed_plates_stats(db_path: str = DEFAULT_DB) -> Dict:
    """
    Получает статистику по выполненным плитам.
    
    Простыми словами:
    - Считает общее количество выполненных плит
    - Считает сколько КП затронуто
    - Возвращает сводку
    
    Возвращает:
        Словарь со статистикой
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        cur = conn.cursor()
        
        # Общее количество записей
        cur.execute('SELECT COUNT(*) FROM completed_plates')
        total_records = cur.fetchone()[0]
        
        # Общее количество плит (сумма qty)
        cur.execute('SELECT SUM(qty) FROM completed_plates')
        result = cur.fetchone()
        total_qty = result[0] if result[0] else 0
        
        # Количество уникальных КП
        cur.execute('SELECT COUNT(DISTINCT kp_id) FROM completed_plates')
        kp_count = cur.fetchone()[0]
        
        # Количество дней производства
        cur.execute('SELECT COUNT(DISTINCT production_day) FROM completed_plates')
        days_count = cur.fetchone()[0]
        
        return {
            'total_records': total_records,
            'total_plates': total_qty,
            'kp_count': kp_count,
            'days_count': days_count
        }
        
    finally:
        conn.close()


def get_completed_plates_by_day(production_day: int, db_path: str = DEFAULT_DB) -> List[Dict]:
    """
    Получает все выполненные плиты за конкретный день производства.
    
    Аргументы:
        production_day: номер дня производства
        db_path: путь к базе данных
    
    Возвращает:
        Список словарей с информацией о плитах
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        
        cur.execute('''
            SELECT cp.*, ko.customer_name
            FROM completed_plates cp
            LEFT JOIN KP_offers ko ON cp.kp_id = ko.kp_id
            WHERE cp.production_day = ?
            ORDER BY cp.kp_id, cp.plate_name
        ''', (production_day,))
        
        return [dict(row) for row in cur.fetchall()]
        
    finally:
        conn.close()


def get_all_plates_in_production(db_path: str = DEFAULT_DB) -> List[Dict]:
    """
    Получает все плиты в производстве (из таблицы kp_plates).
    
    Простыми словами:
    - Возвращает все плиты, которые ещё не выполнены
    - Объединяет с информацией о КП (клиент, дата)
    
    Возвращает:
        Список словарей с информацией о плитах
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        
        cur.execute('''
            SELECT 
                p.id,
                p.kp_id,
                p.position_number,
                p.plate_name,
                p.length_m,
                p.width_m,
                p.load_class,
                p.qty,
                p.unit_weight,
                p.total_weight,
                p.discounted_price,
                p.status as plate_status,
                p.plan_id,
                ko.customer_name,
                ko.execution_terms
            FROM kp_plates p
            LEFT JOIN KP_offers ko ON p.kp_id = ko.kp_id
            LEFT JOIN kp_meta m ON p.kp_id = m.kp_id
            WHERE p.qty > 0 AND p.status = 'в производстве' AND (m.status IS NULL OR m.status = 'в работе')
            ORDER BY p.kp_id, p.position_number
        ''')
        
        return [dict(row) for row in cur.fetchall()]
        
    finally:
        conn.close()


def get_all_completed_plates(db_path: str = DEFAULT_DB) -> List[Dict]:
    """
    Получает все выполненные плиты (из таблицы completed_plates).
    
    Простыми словами:
    - Возвращает все плиты, которые уже выполнены
    - Объединяет с информацией о КП (клиент)
    
    Возвращает:
        Список словарей с информацией о выполненных плитах
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        
        cur.execute('''
            SELECT 
                cp.id,
                cp.kp_id,
                cp.plate_name,
                cp.length_m,
                cp.width_m,
                cp.load_class,
                cp.qty,
                cp.completed_date,
                cp.production_day,
                ko.customer_name,
                ko.execution_terms
            FROM completed_plates cp
            LEFT JOIN KP_offers ko ON cp.kp_id = ko.kp_id
            ORDER BY cp.completed_date DESC, cp.kp_id, cp.plate_name
        ''')
        
        return [dict(row) for row in cur.fetchall()]
        
    finally:
        conn.close()


def get_next_kp_number(db_path: str = DEFAULT_DB) -> int:
    """
    Получает следующий свободный номер КП из базы данных.
    
    Простыми словами:
    - Смотрит максимальный номер КП в таблице KP_offers
    - Возвращает следующий номер (max + 1)
    - Если таблица пуста, возвращает 1
    
    Возвращает:
        Следующий номер КП для нового коммерческого предложения
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        cur = conn.cursor()
        cur.execute('SELECT MAX(kp_id) FROM KP_offers')
        result = cur.fetchone()
        
        max_id = result[0] if result[0] is not None else 0
        next_id = max_id + 1
        
        print(f"[DB] Следующий номер КП: {next_id}")
        return next_id
    
    finally:
        conn.close()


# ==================== ФУНКЦИИ ДЛЯ РАБОТЫ С МЕНЕДЖЕРАМИ ====================

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
    init_schema(db_path)
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
    init_schema(db_path)
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
    init_schema(db_path)
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
    init_schema(db_path)
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
    init_schema(db_path)
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
    init_schema(db_path)
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


def init_default_managers(db_path: str = DEFAULT_DB) -> int:
    """
    Добавляет менеджеров по умолчанию в базу данных.
    
    Простыми словами:
    - Добавляет список менеджеров из вашей таблицы
    - Если менеджер с таким email уже есть, пропускает его
    - Это удобно для первоначальной загрузки данных
    
    Возвращает:
        Количество успешно добавленных менеджеров
    """
    default_managers = [
        ('Булдаков Александр Алексеевич', '79621860029', 'buldakov@gbkstart.ru'),
        ('Зубов Алексей Юрьевич', '79621872265', 'zubov@gbkstart.ru'),
        ('Дургина Ольга Владимировна', '79066395000', 'zakaz@gbkstart.ru'),
        ('Шишов Александр Васильевич', '79206405585', 'shishov@gbkstart.ru'),
        ('Кудигин Никита Валерьевич', '79607428972', 'kuligin@gbkstart.ru'),
    ]
    
    init_schema(db_path)
    added_count = 0
    
    for fio, contact_number, email in default_managers:
        manager_id = add_manager(fio, contact_number, email, db_path)
        if manager_id > 0:
            added_count += 1
    
    print(f"[DB] ✅ Добавлено менеджеров: {added_count} из {len(default_managers)}")
    return added_count


def get_all_kp_list(db_path: str = DEFAULT_DB) -> Dict[str, List[Dict]]:
    """
    Получает все КП, разделенные по статусам.
    
    Простыми словами:
    - Возвращает все КП из базы данных
    - Группирует их по статусам: "в архиве", "в работе", "выполнено"
    - Сортирует по номеру КП (от меньшего к большему)
    
    Возвращает:
        Словарь со списками КП по статусам:
        {
            'archived': [...],      # КП со статусом "в архиве"
            'in_production': [...], # КП со статусом "в работе"
            'completed': [...]      # КП со статусом "выполнено"
        }
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        conn.execute('PRAGMA foreign_keys = ON')
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        
        # Получаем все КП с их статусами
        cur.execute('''
            SELECT 
                ko.kp_id,
                ko.creation_date,
                ko.customer_name,
                ko.manager_name,
                ko.discount_percent,
                ko.subtotal,
                ko.vat_amount,
                ko.total_amount,
                ko.delivery_conditions,
                ko.payment_conditions,
                ko.execution_terms,
                m.status
            FROM KP_offers ko
            LEFT JOIN kp_meta m ON ko.kp_id = m.kp_id
            ORDER BY ko.kp_id ASC
        ''')
        
        all_kp = [dict(row) for row in cur.fetchall()]
        
        # Группируем по статусам
        result = {
            'archived': [],
            'in_production': [],
            'completed': []
        }
        
        for kp in all_kp:
            status = kp.get('status', 'в работе')  # По умолчанию "в работе"
            
            if status == 'в архиве':
                result['archived'].append(kp)
            elif status == 'в работе':
                result['in_production'].append(kp)
            elif status == 'выполнено':
                result['completed'].append(kp)
        
        return result
        
    finally:
        conn.close()


def get_kp_completion_percentage(kp_id: int, db_path: str = DEFAULT_DB) -> Dict:
    """
    Подсчитывает процент выполнения КП.
    
    Простыми словами:
    - Считает сколько плит уже выполнено
    - Считает сколько всего плит в КП
    - Возвращает процент выполнения
    
    Args:
        kp_id: номер КП
        db_path: путь к базе данных
        
    Returns:
        Словарь с информацией:
        {
            'total_plates': int,         # Всего плит
            'completed_plates': int,     # Выполненные плиты  
            'in_production': int,        # В производстве
            'percentage': float          # Процент выполнения (0-100)
        }
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        
        # Считаем плиты в производстве (из kp_plates)
        cur.execute('''
            SELECT COALESCE(SUM(qty), 0) as qty
            FROM kp_plates
            WHERE kp_id = ?
        ''', (kp_id,))
        in_production = cur.fetchone()['qty']
        
        # Считаем выполненные плиты (из completed_plates)
        cur.execute('''
            SELECT COALESCE(SUM(qty), 0) as qty
            FROM completed_plates
            WHERE kp_id = ?
        ''', (kp_id,))
        completed = cur.fetchone()['qty']
        
        # Всего плит
        total = in_production + completed
        
        # Процент выполнения
        if total > 0:
            percentage = (completed / total) * 100
        else:
            percentage = 0.0
        
        return {
            'total_plates': total,
            'completed_plates': completed,
            'in_production': in_production,
            'percentage': round(percentage, 1)
        }
        
    finally:
        conn.close()


def get_kp_plates_in_plan_percentage(kp_id: int, db_path: str = DEFAULT_DB) -> Dict:
    """
    Подсчитывает, какой процент плит КП уже в плане производства.

    Считает сумму qty по kp_plates со статусом 'в производстве' или 'в плане' (база),
    и сумму qty со статусом 'в плане'; возвращает процент (in_plan / total * 100).

    Args:
        kp_id: номер КП
        db_path: путь к базе данных

    Returns:
        Словарь: 'total_plates', 'in_plan', 'percentage' (0-100)
    """
    init_schema(db_path)
    conn = _connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute('''
            SELECT COALESCE(SUM(qty), 0) as total
            FROM kp_plates
            WHERE kp_id = ? AND status IN ('в производстве', 'в плане')
        ''', (kp_id,))
        total = cur.fetchone()[0]
        cur.execute('''
            SELECT COALESCE(SUM(qty), 0) as in_plan
            FROM kp_plates
            WHERE kp_id = ? AND status = 'в плане'
        ''', (kp_id,))
        in_plan = cur.fetchone()[0]
        percentage = (in_plan / total * 100) if total > 0 else 0.0
        return {
            'total_plates': total,
            'in_plan': in_plan,
            'percentage': round(percentage, 1)
        }
    finally:
        conn.close()


def get_kp_total_length(kp_id: int, db_path: str = DEFAULT_DB) -> float:
    """
    Суммарная длина плит КП в метрах: SUM(COALESCE(length_m, 0) * qty)
    по строкам kp_plates со статусом 'в производстве' или 'в плане'.
    """
    init_schema(db_path)
    conn = _connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute('''
            SELECT COALESCE(SUM(COALESCE(length_m, 0) * qty), 0.0)
            FROM kp_plates
            WHERE kp_id = ? AND status IN ('в производстве', 'в плане')
        ''', (kp_id,))
        return float(cur.fetchone()[0])
    finally:
        conn.close()


def update_kp_execution_date(kp_id: int, new_date: str, db_path: str = DEFAULT_DB) -> bool:
    """
    Обновляет дату выполнения для КП.
    
    Простыми словами:
    - Меняет срок выполнения для КП
    - Это влияет на все плиты в этом КП
    
    Args:
        kp_id: номер КП
        new_date: новая дата в формате "DD.MM.YYYY"
        db_path: путь к базе данных
        
    Returns:
        True при успехе, False при ошибке
    """
    init_schema(db_path)
    conn = _connect(db_path)
    
    try:
        cur = conn.cursor()
        
        # Обновляем дату выполнения
        cur.execute('''
            UPDATE KP_offers 
            SET execution_terms = ? 
            WHERE kp_id = ?
        ''', (new_date, kp_id))
        
        conn.commit()
        
        # Проверяем, что обновление произошло
        return cur.rowcount > 0
        
    except Exception as e:
        conn.rollback()
        return False
        
    finally:
        conn.close()


# ==================== ПРИМЕР ИСПОЛЬЗОВАНИЯ ====================

if __name__ == '__main__':
    # Пример тестовых данных
    test_order_data = [
        {
            'name': 'ПБ 78-12-8п',
            'length_m': 7.8,
            'width_m': 1.2,
            'qty': 4,
            'load_class': 800,
            'unit_price': 5000.0,
            'weight': 987.84
        },
        {
            'name': 'ПБ 66-3-8п',
            'length_m': 6.6,
            'width_m': 0.3,
            'qty': 2,
            'load_class': 800,
            'unit_price': 1200.0,
            'weight': 104.54
        }
    ]
    
    print("📝 Создаю тестовую БД...")
    
    # Пример сохранения КП в БД
    kp_id = save_kp_to_db(
        creation_date='01.01.2024',
        order_data=test_order_data,
        xlsx_file_path=None,  # Можно указать путь к файлу
        customer_name='ООО Тест',
        manager_name='Иванов И.И.',
        discount_percent=5.0,
        delivery_conditions='Самовывоз',
        payment_conditions='Предоплата 100%',
        execution_terms='14 дней',
        status='в работе'
    )
    print(f"✅ КП сохранён в БД с порядковым номером: {kp_id}")
    
    # Пример получения КП
    print("\n🔍 Ищу КП по номеру...")
    kp = get_kp_by_id(kp_id)
    if kp:
        print(f"✅ Найден КП № {kp['kp_id']}")
        print(f"   Дата: {kp['creation_date']}")
        print(f"   Клиент: {kp['customer_name']}")
        print(f"   Менеджер: {kp['manager_name']}")
        print(f"   Сумма: {kp['total_amount']:.2f} ₽")
        print(f"   Статус: {kp.get('status', 'не указан')}")
        print(f"   Плит в заказе: {len(kp['plates'])} позиций")
    
    # Пример получения всех КП в работе
    print("\n📋 Все КП в работе:")
    all_kp = get_all_kp_by_status('в работе')
    for kp in all_kp:
        print(f"  • КП № {kp['kp_id']} - {kp['customer_name']} - {kp['total_amount']:.2f} ₽")
