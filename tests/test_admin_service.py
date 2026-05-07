from __future__ import annotations

import json
import sqlite3
from pathlib import Path
from unittest.mock import MagicMock

import pytest

from app.core.settings import Settings
from app.repositories.auth_repository import AuthRepository
from app.services.admin_service import AdminService
from core import kp_db


@pytest.fixture()
def admin_settings(tmp_path: Path) -> Settings:
    """Изолированный Settings, указывающий на директории внутри tmp_path."""
    plans_dir = tmp_path / "plans"
    plans_dir.mkdir()
    settings = Settings(
        plita_db_path=tmp_path / "plita.db",
        plans_dir=plans_dir,
        plans_metadata_path=tmp_path / "plans_metadata.json",
        current_plan_path=tmp_path / "current_plan.json",
        work_calendar_path=tmp_path / "work_calendar.json",
    )
    return settings


@pytest.fixture()
def populated_db(admin_settings: Settings) -> Path:
    """Заполненная плитами и КП тестовая база + таблица app_users с админом."""
    db_path = str(admin_settings.plita_db_path)
    kp_db.init_schema(db_path)

    auth_repo = AuthRepository(db_path=db_path)
    auth_repo.create_or_update_user(
        username="root_admin",
        password="qwerty123",
        role="admin",
    )

    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO KP_offers (creation_date, customer_name) VALUES (?, ?)",
            ("01.01.2026", "ООО Тест"),
        )
        kp_id = cur.lastrowid
        cur.execute(
            "INSERT INTO kp_meta (kp_id, status) VALUES (?, ?)",
            (kp_id, "в работе"),
        )
        cur.execute(
            """
            INSERT INTO kp_plates (kp_id, position_number, plate_name, length_m, width_m, qty)
            VALUES (?, 1, 'ПБ 78-12-8п', 7.8, 1.2, 2)
            """,
            (kp_id,),
        )
        cur.execute(
            """
            INSERT INTO completed_plates (kp_id, plate_name, length_m, width_m, qty, completed_date)
            VALUES (?, 'ПБ 78-12-8п', 7.8, 1.2, 1, '01.02.2026')
            """,
            (kp_id,),
        )
        cur.execute(
            """
            INSERT INTO plate_rests
                (kp_id, source_plate_name, rest_width_mm, length_m, qty, created_date)
            VALUES (?, 'ПБ 78-12-8п', 200, 1.5, 1, '01.02.2026')
            """,
            (kp_id,),
        )
        cur.execute(
            "INSERT INTO kp_files (kp_id, file_path) VALUES (?, '/tmp/x.xlsx')",
            (kp_id,),
        )
        cur.execute(
            """
            INSERT INTO plate_status_log (
                plate_id, kp_id, plate_name, plan_id, day_number,
                from_status, to_status, qty, reason
            )
            VALUES (1, ?, 'ПБ 78-12-8п', 'plan_test', 1,
                    'в плане', 'completed', 1, 'completed')
            """,
            (kp_id,),
        )
        conn.commit()
    return Path(db_path)


@pytest.fixture()
def populated_plans(admin_settings: Settings) -> None:
    """Создаёт JSON-файлы планов и календарь для проверки очистки."""
    admin_settings.plans_metadata_path.write_text(
        json.dumps({"plans": [{"plan_id": "p1"}]}), encoding="utf-8"
    )
    admin_settings.current_plan_path.write_text(
        json.dumps({"id": "current"}), encoding="utf-8"
    )
    (admin_settings.plans_dir / "plan_1.json").write_text(
        json.dumps({"id": "plan_1"}), encoding="utf-8"
    )
    (admin_settings.plans_dir / "plan_2.json").write_text(
        json.dumps({"id": "plan_2"}), encoding="utf-8"
    )
    admin_settings.work_calendar_path.write_text(
        json.dumps({"extra_holidays": ["2026-01-01"], "extra_workdays": []}),
        encoding="utf-8",
    )


def _make_service(
    admin_settings: Settings,
    *,
    plan_repository: MagicMock | None = None,
) -> AdminService:
    if plan_repository is None:
        plan_repository = MagicMock()
        plan_repository.list_metadata.return_value = {"plans": []}
    return AdminService(
        settings=admin_settings,
        plan_repository=plan_repository,
    )


def _table_count(db_path: Path | str, table: str) -> int:
    with sqlite3.connect(str(db_path)) as conn:
        cur = conn.cursor()
        cur.execute(f"SELECT COUNT(*) FROM {table}")
        return int(cur.fetchone()[0])


def test_reset_full_clears_all_plate_tables_and_plans(
    admin_settings: Settings,
    populated_db: Path,
    populated_plans: None,
) -> None:
    service = _make_service(admin_settings)

    report = service.reset_full()

    assert _table_count(populated_db, "KP_offers") == 0
    assert _table_count(populated_db, "kp_plates") == 0
    assert _table_count(populated_db, "completed_plates") == 0
    assert _table_count(populated_db, "plate_rests") == 0
    assert _table_count(populated_db, "kp_files") == 0
    assert _table_count(populated_db, "kp_meta") == 0
    assert _table_count(populated_db, "plate_status_log") == 0

    assert not admin_settings.current_plan_path.exists()
    assert not admin_settings.plans_metadata_path.exists()
    assert admin_settings.plans_dir.exists()
    assert list(admin_settings.plans_dir.glob("*.json")) == []

    assert report.plans["plan_files"] == 2
    assert report.plans["current_plan"] == 1
    assert report.plans["metadata"] == 1
    assert report.calendar_reset is True

    calendar_data = json.loads(admin_settings.work_calendar_path.read_text(encoding="utf-8"))
    assert calendar_data == {"extra_holidays": [], "extra_workdays": []}


def test_reset_full_preserves_app_users(
    admin_settings: Settings,
    populated_db: Path,
    populated_plans: None,
) -> None:
    """Самое главное: админ не должен потерять собственный аккаунт."""
    service = _make_service(admin_settings)

    service.reset_full()

    assert _table_count(populated_db, "app_users") == 1
    auth_repo = AuthRepository(db_path=str(populated_db))
    user = auth_repo.get_user_by_username("root_admin")
    assert user is not None
    assert user["role"] == "admin"


def test_reset_kp_only_keeps_completed_plates_and_rests(
    admin_settings: Settings,
    populated_db: Path,
) -> None:
    """`reset_kp_only` НЕ должна очищать таблицы completed_plates и plate_rests
    напрямую (только косвенно через ON DELETE CASCADE). Висячие записи без
    родительской KP_offers должны остаться."""
    with sqlite3.connect(str(populated_db)) as conn:
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO completed_plates
                (kp_id, plate_name, length_m, width_m, qty, completed_date)
            VALUES (9999, 'Orphan', 1.0, 1.0, 1, '01.02.2026')
            """
        )
        cur.execute(
            """
            INSERT INTO plate_rests
                (kp_id, source_plate_name, rest_width_mm, length_m, qty, created_date)
            VALUES (9999, 'Orphan', 100, 1.0, 1, '01.02.2026')
            """
        )
        conn.commit()

    service = _make_service(admin_settings)

    service.reset_kp_only()

    assert _table_count(populated_db, "KP_offers") == 0
    assert _table_count(populated_db, "kp_plates") == 0
    assert _table_count(populated_db, "kp_files") == 0
    assert _table_count(populated_db, "kp_meta") == 0
    assert _table_count(populated_db, "completed_plates") == 1
    assert _table_count(populated_db, "plate_rests") == 1


def test_reset_plans_only_does_not_touch_db(
    admin_settings: Settings,
    populated_db: Path,
    populated_plans: None,
) -> None:
    service = _make_service(admin_settings)

    report = service.reset_plans_only()

    assert _table_count(populated_db, "KP_offers") == 1
    assert _table_count(populated_db, "kp_plates") == 1
    assert report.plans["plan_files"] == 2
    assert not admin_settings.current_plan_path.exists()
    assert not admin_settings.plans_metadata_path.exists()


def test_reset_calendar_only_writes_empty_calendar(admin_settings: Settings) -> None:
    admin_settings.work_calendar_path.write_text(
        json.dumps({"extra_holidays": ["2026-01-01"], "extra_workdays": []}),
        encoding="utf-8",
    )
    service = _make_service(admin_settings)

    report = service.reset_calendar_only()

    assert report.calendar_reset is True
    data = json.loads(admin_settings.work_calendar_path.read_text(encoding="utf-8"))
    assert data == {"extra_holidays": [], "extra_workdays": []}


def test_get_stats_aggregates_db_and_plans(
    admin_settings: Settings,
    populated_db: Path,
    populated_plans: None,
) -> None:
    plan_repo = MagicMock()
    plan_repo.list_metadata.return_value = {
        "plans": [{"plan_id": "p1"}, {"plan_id": "p2"}, {"plan_id": "p3"}]
    }
    service = _make_service(admin_settings, plan_repository=plan_repo)

    stats = service.get_stats()

    assert stats.kp_total == 1
    assert stats.kp_in_work == 1
    assert stats.kp_completed == 0
    assert stats.plates_in_work == 1
    assert stats.plates_completed == 1
    assert stats.plate_rests == 1
    assert stats.plans_count == 3
    assert stats.current_plan_present is True
