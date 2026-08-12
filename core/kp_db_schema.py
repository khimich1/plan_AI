#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""SQLite schema initialization for plita.db (A1 / A4)."""

from __future__ import annotations

import os
import sqlite3
import threading

from core.kp_db_common import DEFAULT_DB, _connect

_schema_ready: set[str] = set()
_schema_lock = threading.Lock()

def _init_schema_impl(db_path: str = DEFAULT_DB) -> None:
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
                line_id TEXT,
                FOREIGN KEY (kp_id) REFERENCES KP_offers(kp_id) ON DELETE CASCADE
            )
        ''')
        
        # Миграция: колонка unit_price для пересчёта скидки и генерации документов из архива
        try:
            cur.execute("ALTER TABLE kp_plates ADD COLUMN unit_price REAL")
            conn.commit()
        except sqlite3.OperationalError:
            pass  # колонка уже существует

        try:
            cur.execute("ALTER TABLE kp_plates ADD COLUMN concrete_grade TEXT")
            conn.commit()
        except sqlite3.OperationalError:
            pass

        # Стоимость одного рейса (логистика), как в draft metadata logistics_cost / PDF/XLSX
        try:
            cur.execute("ALTER TABLE KP_offers ADD COLUMN logistics_cost REAL DEFAULT 0")
            conn.commit()
        except sqlite3.OperationalError:
            pass
        
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
        # status - статус КП: выполнено, отклонено, в работе, в ожидании, На СГП
        # ordered_qty - снимок заказного qty (M для бейджа N/M); freeze при уходе в производство
        cur.execute('''
            CREATE TABLE IF NOT EXISTS kp_meta (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kp_id INTEGER NOT NULL,
                status TEXT DEFAULT 'в работе',
                ordered_qty INTEGER,
                FOREIGN KEY (kp_id) REFERENCES KP_offers(kp_id) ON DELETE CASCADE,
                UNIQUE(kp_id)
            )
        ''')
        
        # Таблица 5: completed_plates — склад готовой продукции (СГП)
        # kp_id NULLABLE: отвязанные плиты остаются на складе без КП
        # plan_id — план, с которого плита пришла; при удалении плана обнуляется
        cur.execute('''
            CREATE TABLE IF NOT EXISTS completed_plates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kp_id INTEGER,
                plate_name TEXT NOT NULL,
                length_m REAL,
                width_m REAL,
                load_class INTEGER,
                qty INTEGER NOT NULL,
                completed_date TEXT NOT NULL,
                production_day INTEGER,
                nomenclature_id TEXT,
                plan_id TEXT,
                FOREIGN KEY (kp_id) REFERENCES KP_offers(kp_id) ON DELETE SET NULL
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

        if 'length_dm_raw' not in columns:
            print("[DB] Миграция: добавляем колонку length_dm_raw в kp_plates...")
            cur.execute("ALTER TABLE kp_plates ADD COLUMN length_dm_raw TEXT")
            print("[DB] ✅ Колонка length_dm_raw добавлена")

        cur.execute(
            'CREATE INDEX IF NOT EXISTS idx_plates_plan_day '
            'ON kp_plates(plan_id, day_number)'
        )
        cur.execute(
            'CREATE INDEX IF NOT EXISTS idx_plates_kp_length_dm '
            'ON kp_plates(kp_id, length_dm_raw)'
        )
        cur.execute(
            'CREATE INDEX IF NOT EXISTS idx_plates_kp_plate_name '
            'ON kp_plates(kp_id, plate_name)'
        )

        # Таблица 9: production_plans — производственные планы (A2 / WP3).
        # План хранится как JSON в payload_json; version — optimistic concurrency.
        cur.execute('''
            CREATE TABLE IF NOT EXISTS production_plans (
                id TEXT PRIMARY KEY,
                payload_json TEXT NOT NULL,
                version INTEGER NOT NULL DEFAULT 1,
                is_active INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL DEFAULT (datetime('now')),
                updated_at TEXT NOT NULL DEFAULT (datetime('now'))
            )
        ''')
        cur.execute(
            'CREATE INDEX IF NOT EXISTS idx_production_plans_active '
            'ON production_plans(is_active)'
        )
        
        # Обратная засылка length_dm_raw из plate_name для существующих строк (один раз после ADD COLUMN)
        # length_m may be absent on very old shapes — use NULL fallback then.
        length_m_expr = "length_m" if "length_m" in columns else "NULL"
        cur.execute(
            f"SELECT id, plate_name, {length_m_expr} FROM kp_plates "
            "WHERE length_dm_raw IS NULL OR length_dm_raw = ''"
        )
        rows = cur.fetchall()
        if rows:
            from core.config_and_data import extract_length_dm_raw_from_plate_name
            for row in rows:
                rid, plate_name, length_m = row
                raw = extract_length_dm_raw_from_plate_name(plate_name or "")
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

        if "owner_user_id" not in meta_columns:
            print("[DB] Миграция: добавляем колонку owner_user_id в kp_meta...")
            cur.execute("ALTER TABLE kp_meta ADD COLUMN owner_user_id INTEGER")
            print("[DB] ✅ Колонка owner_user_id добавлена в kp_meta")

        if "ordered_qty" not in meta_columns:
            print("[DB] Миграция: добавляем колонку ordered_qty в kp_meta...")
            cur.execute("ALTER TABLE kp_meta ADD COLUMN ordered_qty INTEGER")
            print("[DB] ✅ Колонка ordered_qty добавлена в kp_meta")

        if "product_type" not in meta_columns:
            print("[DB] Миграция: добавляем колонку product_type в kp_meta...")
            cur.execute(
                "ALTER TABLE kp_meta ADD COLUMN product_type TEXT DEFAULT 'plates'"
            )
            cur.execute(
                "UPDATE kp_meta SET product_type = 'plates' WHERE product_type IS NULL"
            )
            print("[DB] ✅ Колонка product_type добавлена в kp_meta")

        # Таблица kp_piles — позиции КП на сваи (отдельно от kp_plates)
        cur.execute('''
            CREATE TABLE IF NOT EXISTS kp_piles (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kp_id INTEGER NOT NULL,
                position_number INTEGER NOT NULL,
                mark TEXT NOT NULL,
                concrete_grade TEXT NOT NULL,
                qty INTEGER NOT NULL,
                unit_price REAL NOT NULL,
                discounted_price REAL NOT NULL,
                line_id TEXT,
                FOREIGN KEY (kp_id) REFERENCES KP_offers(kp_id) ON DELETE CASCADE
            )
        ''')
        cur.execute(
            'CREATE INDEX IF NOT EXISTS idx_kp_id_piles ON kp_piles(kp_id)'
        )

        # Таблица kp_steps — позиции КП на лестничные ступени (без класса бетона)
        cur.execute('''
            CREATE TABLE IF NOT EXISTS kp_steps (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kp_id INTEGER NOT NULL,
                position_number INTEGER NOT NULL,
                mark TEXT NOT NULL,
                qty INTEGER NOT NULL,
                unit_price REAL NOT NULL,
                discounted_price REAL NOT NULL,
                line_id TEXT,
                FOREIGN KEY (kp_id) REFERENCES KP_offers(kp_id) ON DELETE CASCADE
            )
        ''')
        cur.execute(
            'CREATE INDEX IF NOT EXISTS idx_kp_id_steps ON kp_steps(kp_id)'
        )

        # Таблица kp_marches — позиции КП на лестничные марши (с классом бетона)
        cur.execute('''
            CREATE TABLE IF NOT EXISTS kp_marches (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kp_id INTEGER NOT NULL,
                position_number INTEGER NOT NULL,
                mark TEXT NOT NULL,
                concrete_grade TEXT NOT NULL,
                qty INTEGER NOT NULL,
                unit_price REAL NOT NULL,
                discounted_price REAL NOT NULL,
                line_id TEXT,
                FOREIGN KEY (kp_id) REFERENCES KP_offers(kp_id) ON DELETE CASCADE
            )
        ''')
        cur.execute(
            'CREATE INDEX IF NOT EXISTS idx_kp_id_marches ON kp_marches(kp_id)'
        )

        # Таблица kp_bridge_piles — позиции КП на мостовые сваи (отдельно от kp_piles)
        cur.execute('''
            CREATE TABLE IF NOT EXISTS kp_bridge_piles (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kp_id INTEGER NOT NULL,
                position_number INTEGER NOT NULL,
                mark TEXT NOT NULL,
                concrete_grade TEXT NOT NULL,
                qty INTEGER NOT NULL,
                unit_price REAL NOT NULL,
                discounted_price REAL NOT NULL,
                line_id TEXT,
                FOREIGN KEY (kp_id) REFERENCES KP_offers(kp_id) ON DELETE CASCADE
            )
        ''')
        cur.execute(
            'CREATE INDEX IF NOT EXISTS idx_kp_id_bridge_piles ON kp_bridge_piles(kp_id)'
        )

        # Таблица kp_fbs — позиции КП на ФБС (отдельно от kp_piles / kp_bridge_piles)
        cur.execute('''
            CREATE TABLE IF NOT EXISTS kp_fbs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kp_id INTEGER NOT NULL,
                position_number INTEGER NOT NULL,
                mark TEXT NOT NULL,
                concrete_grade TEXT NOT NULL,
                qty INTEGER NOT NULL,
                unit_price REAL NOT NULL,
                discounted_price REAL NOT NULL,
                line_id TEXT,
                FOREIGN KEY (kp_id) REFERENCES KP_offers(kp_id) ON DELETE CASCADE
            )
        ''')
        cur.execute(
            'CREATE INDEX IF NOT EXISTS idx_kp_id_fbs ON kp_fbs(kp_id)'
        )

        # MNA-301: line_id на всех kp_* line-таблицах (idempotent ALTER для legacy БД).
        # kp_plates уже мигрирована выше (try/except рядом с unit_price); остальные — здесь.
        _ensure_line_id_columns(cur)

        # === МИГРАЦИЯ: Добавляем nomenclature_id ===
        if 'nomenclature_id' not in columns:
            print("[DB] Миграция: добавляем колонку nomenclature_id в kp_plates...")
            cur.execute("ALTER TABLE kp_plates ADD COLUMN nomenclature_id TEXT")
            print("[DB] Колонка nomenclature_id добавлена в kp_plates")

        _migrate_completed_plates_for_sgp(cur)
        cur.execute(
            "CREATE INDEX IF NOT EXISTS idx_completed_plan_id ON completed_plates(plan_id)"
        )

        _init_shipment_logistics_schema(cur)

        conn.commit()
    finally:
        conn.close()


def _ensure_line_id_columns(cur: sqlite3.Cursor) -> None:
    """MNA-301: add nullable ``line_id TEXT`` to every kp_* line table (idempotent).

    Fresh CREATE already includes the column; this covers legacy DBs where
    ``CREATE TABLE IF NOT EXISTS`` left the old shape unchanged.
    ``kp_meta.product_type`` remains unconstrained TEXT (accepts ``mixed``).
    """
    line_tables = (
        "kp_plates",
        "kp_piles",
        "kp_steps",
        "kp_marches",
        "kp_bridge_piles",
        "kp_fbs",
    )
    for table in line_tables:
        cur.execute(f"PRAGMA table_info({table})")
        columns = {row[1] for row in cur.fetchall()}
        if "line_id" in columns:
            continue
        print(f"[DB] Миграция: добавляем колонку line_id в {table}...")
        cur.execute(f"ALTER TABLE {table} ADD COLUMN line_id TEXT")
        print(f"[DB] ✅ Колонка line_id добавлена в {table}")


def _init_shipment_logistics_schema(cur: sqlite3.Cursor) -> None:
    """Таблицы раздела «Логистика» (SHIP-000): рейсы, состав, справочники.

    ``shipment_orders.kp_id`` — NULLABLE + ON DELETE SET NULL: рейс переживает
    удаление КП (план P-H). ``shipment_items.completed_plate_id``/``kp_id`` —
    snapshot-ссылки: списанные плиты остаются в completed_plates с qty=0,
    поэтому done-рейс всегда можно показать.
    """
    cur.execute('''
        CREATE TABLE IF NOT EXISTS shipments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            shipment_date TEXT NOT NULL,
            delivery_type TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'in_work',
            attention INTEGER NOT NULL DEFAULT 0,
            attention_comment TEXT,
            carrier_id INTEGER REFERENCES carriers(id),
            driver_name TEXT,
            vehicle_text TEXT,
            vehicle_class TEXT,
            proxy_no TEXT,
            upd_no TEXT,
            freight_request_no TEXT,
            planned_cost REAL,
            time_slot TEXT,
            propose_snapshot TEXT,
            completed_at TEXT,
            actor TEXT,
            created_at TEXT NOT NULL DEFAULT (datetime('now'))
        )
    ''')
    cur.execute('''
        CREATE TABLE IF NOT EXISTS shipment_orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            shipment_id INTEGER NOT NULL REFERENCES shipments(id) ON DELETE CASCADE,
            kp_id INTEGER REFERENCES KP_offers(kp_id) ON DELETE SET NULL,
            ya_order_no TEXT,
            UNIQUE (shipment_id, kp_id)
        )
    ''')
    cur.execute('''
        CREATE TABLE IF NOT EXISTS shipment_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            shipment_id INTEGER NOT NULL REFERENCES shipments(id) ON DELETE CASCADE,
            item_type TEXT NOT NULL,
            completed_plate_id INTEGER REFERENCES completed_plates(id),
            kp_id INTEGER,
            mark TEXT,
            qty INTEGER NOT NULL,
            unit_weight_kg REAL,
            weight_kg REAL,
            sort_order INTEGER NOT NULL DEFAULT 0,
            note TEXT
        )
    ''')
    cur.execute('''
        CREATE TABLE IF NOT EXISTS carriers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            name_normalized TEXT NOT NULL,
            source_sheet TEXT,
            note TEXT,
            active INTEGER NOT NULL DEFAULT 1,
            merged_into_id INTEGER REFERENCES carriers(id),
            created_at TEXT NOT NULL DEFAULT (datetime('now'))
        )
    ''')
    cur.execute('''
        CREATE TABLE IF NOT EXISTS pile_catalog (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            mark TEXT NOT NULL UNIQUE,
            length_m REAL,
            section_mm INTEGER,
            volume_m3 REAL,
            weight_kg REAL NOT NULL,
            pcs_per_20t INTEGER
        )
    ''')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_shipments_date ON shipments(shipment_date)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_shipments_status ON shipments(status)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_shipments_carrier ON shipments(carrier_id)')
    cur.execute(
        'CREATE INDEX IF NOT EXISTS idx_shipment_orders_shipment ON shipment_orders(shipment_id)'
    )
    cur.execute('CREATE INDEX IF NOT EXISTS idx_shipment_orders_kp ON shipment_orders(kp_id)')
    cur.execute(
        'CREATE INDEX IF NOT EXISTS idx_shipment_items_shipment ON shipment_items(shipment_id)'
    )
    cur.execute(
        'CREATE INDEX IF NOT EXISTS idx_shipment_items_cp ON shipment_items(completed_plate_id)'
    )
    cur.execute('CREATE INDEX IF NOT EXISTS idx_carriers_active ON carriers(active)')
    # Один активный перевозчик на нормализованное имя; слитые дубли слот освобождают.
    cur.execute(
        '''
        CREATE UNIQUE INDEX IF NOT EXISTS idx_carriers_name_normalized_active
        ON carriers(name_normalized) WHERE merged_into_id IS NULL
        '''
    )

    cur.execute("PRAGMA table_info(plate_status_log)")
    log_columns = {row[1] for row in cur.fetchall()}
    if "shipment_id" not in log_columns:
        print("[DB] Миграция: добавляем колонку shipment_id в plate_status_log...")
        cur.execute("ALTER TABLE plate_status_log ADD COLUMN shipment_id INTEGER")
        print("[DB] ✅ Колонка shipment_id добавлена в plate_status_log")
    cur.execute(
        'CREATE INDEX IF NOT EXISTS idx_status_log_shipment ON plate_status_log(shipment_id)'
    )


def _migrate_completed_plates_for_sgp(cur: sqlite3.Cursor) -> None:
    """Make ``completed_plates.kp_id`` nullable + add ``plan_id`` (SGP-000).

    SQLite cannot ALTER column nullability / FK action, so when ``kp_id`` is
    still NOT NULL we rebuild the table (copy → drop → rename).
    """
    cur.execute("PRAGMA table_info(completed_plates)")
    cp_rows = cur.fetchall()
    if not cp_rows:
        return

    cp_by_name = {row[1]: row for row in cp_rows}
    kp_notnull = int(cp_by_name.get("kp_id", (None, None, None, 1))[3] or 0)
    has_plan_id = "plan_id" in cp_by_name
    has_nomenclature = "nomenclature_id" in cp_by_name

    if not has_nomenclature:
        print("[DB] Миграция: добавляем колонку nomenclature_id в completed_plates...")
        cur.execute("ALTER TABLE completed_plates ADD COLUMN nomenclature_id TEXT")
        print("[DB] Колонка nomenclature_id добавлена в completed_plates")
        has_nomenclature = True

    if kp_notnull == 0 and has_plan_id:
        return

    if kp_notnull == 0 and not has_plan_id:
        print("[DB] Миграция: добавляем колонку plan_id в completed_plates...")
        cur.execute("ALTER TABLE completed_plates ADD COLUMN plan_id TEXT")
        print("[DB] ✅ Колонка plan_id добавлена в completed_plates")
        return

    print("[DB] Миграция СГП: пересоздаём completed_plates (kp_id NULLABLE, plan_id)...")
    cur.execute("PRAGMA foreign_keys = OFF")
    cur.execute(
        """
        CREATE TABLE completed_plates_sgp (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            kp_id INTEGER,
            plate_name TEXT NOT NULL,
            length_m REAL,
            width_m REAL,
            load_class INTEGER,
            qty INTEGER NOT NULL,
            completed_date TEXT NOT NULL,
            production_day INTEGER,
            nomenclature_id TEXT,
            plan_id TEXT,
            FOREIGN KEY (kp_id) REFERENCES KP_offers(kp_id) ON DELETE SET NULL
        )
        """
    )
    nomenclature_expr = "nomenclature_id" if has_nomenclature else "NULL"
    plan_expr = "plan_id" if has_plan_id else "NULL"
    cur.execute(
        f"""
        INSERT INTO completed_plates_sgp (
            id, kp_id, plate_name, length_m, width_m, load_class,
            qty, completed_date, production_day, nomenclature_id, plan_id
        )
        SELECT
            id, kp_id, plate_name, length_m, width_m, load_class,
            qty, completed_date, production_day, {nomenclature_expr}, {plan_expr}
        FROM completed_plates
        """
    )
    cur.execute("DROP TABLE completed_plates")
    cur.execute("ALTER TABLE completed_plates_sgp RENAME TO completed_plates")
    cur.execute(
        "CREATE INDEX IF NOT EXISTS idx_completed_kp_id ON completed_plates(kp_id)"
    )
    cur.execute(
        "CREATE INDEX IF NOT EXISTS idx_completed_date ON completed_plates(completed_date)"
    )
    cur.execute(
        "CREATE INDEX IF NOT EXISTS idx_completed_plan_id ON completed_plates(plan_id)"
    )
    cur.execute("PRAGMA foreign_keys = ON")
    print("[DB] ✅ completed_plates мигрирована под СГП")

def ensure_schema(db_path: str = DEFAULT_DB) -> None:
    """Idempotent schema initialization (once per absolute db path per process)."""
    abs_path = os.path.abspath(db_path)
    if abs_path in _schema_ready:
        return
    with _schema_lock:
        if abs_path in _schema_ready:
            return
        _init_schema_impl(db_path)
        _schema_ready.add(abs_path)


def init_schema(db_path: str = DEFAULT_DB) -> None:
    """Backward-compatible alias for :func:`ensure_schema`."""
    ensure_schema(db_path)
