"""Сервис администрирования базы данных и хранилища планов.

Инкапсулирует «опасные» операции:

- получение сводной статистики;
- полное обнуление (SQLite КП/плиты + SQLite-планы + legacy JSON + календарь);
- частичные обнуления (только КП / только планы / только календарь);
- восстановление «застрявших» плит.

Все мутации идут только через этот сервис, чтобы FastAPI-endpoint'ы
оставались тонкими, а пути к файлам брались из ``Settings``.

ВНИМАНИЕ: деструктивные операции требуют ``require_destructive_db_reset``;
при параллельной записи возможна гонка за ``data/plans/*`` и ``plita.db``.
"""
from __future__ import annotations

import logging
import shutil
from pathlib import Path

from app.core.settings import Settings, get_settings
from app.repositories.plan_repository import PlanRepository
from app.repositories.work_calendar_repository import WorkCalendarRepository
from app.schemas.admin import DbResetReport, DbStatsResponse, RecoverPlatesResponse
from core import kp_db_offers, kp_db_plates
from core.destructive_db_guard import require_destructive_db_reset

logger = logging.getLogger(__name__)

EMPTY_CALENDAR: dict[str, list[str]] = {"extra_holidays": [], "extra_workdays": []}


class AdminService:
    """Высокоуровневые операции администрирования БД."""

    def __init__(
        self,
        settings: Settings | None = None,
        plan_repository: PlanRepository | None = None,
        calendar_repository: WorkCalendarRepository | None = None,
    ) -> None:
        self.settings = settings or get_settings()
        self.plan_repository = plan_repository or PlanRepository()
        self.calendar_repository = calendar_repository or WorkCalendarRepository(
            settings=self.settings
        )
        self.db_path = str(self.settings.plita_db_path)

    # ---------- Stats ----------

    def get_stats(self) -> DbStatsResponse:
        sqlite_stats = kp_db_offers.get_db_stats(self.db_path)
        plans_metadata = self.plan_repository.list_metadata()
        plans = plans_metadata.get("plans") if isinstance(plans_metadata, dict) else None
        plans_count = len(plans) if isinstance(plans, list) else 0
        return DbStatsResponse(
            kp_total=int(sqlite_stats.get("kp_total", 0)),
            kp_in_work=int(sqlite_stats.get("kp_in_work", 0)),
            kp_completed=int(sqlite_stats.get("kp_completed", 0)),
            plates_in_work=int(sqlite_stats.get("plates_in_work", 0)),
            plates_completed=int(sqlite_stats.get("plates_completed", 0)),
            plate_rests=int(sqlite_stats.get("plate_rests", 0)),
            plans_count=plans_count,
            legacy_json_files_count=self._count_legacy_json_files(
                self.settings.plans_dir,
                self.settings.archived_data_dir / "plans",
            ),
            current_plan_present=self.settings.current_plan_path.exists(),
        )

    # ---------- Reset ----------

    def reset_full(self) -> DbResetReport:
        """Полное обнуление: SQLite КП/плиты + SQLite-планы + legacy JSON + календарь.

        Таблицы ``app_users`` и ``managers`` НЕ затрагиваются — администратор
        не теряет сессию.
        """
        require_destructive_db_reset()
        sqlite_report = kp_db_offers.clear_all_plates_data(self.db_path)
        plans_report = self._clear_all_plans()
        plans_report.update(self._clear_archived_legacy())
        calendar_reset = self._reset_calendar()
        return DbResetReport(
            sqlite=_normalize_int_dict(sqlite_report),
            plans=plans_report,
            calendar_reset=calendar_reset,
        )

    def reset_kp_only(self) -> DbResetReport:
        """Частичное обнуление: только таблицы КП.

        НЕ трогает ``completed_plates`` и ``plate_rests``.
        Соответствует старой команде бота ``/clear_all_kp``.
        """
        require_destructive_db_reset()
        sqlite_report = kp_db_offers.clear_all_kp(self.db_path)
        return DbResetReport(sqlite=_normalize_int_dict(sqlite_report))

    def reset_plans_only(self) -> DbResetReport:
        """Частичное обнуление: SQLite ``production_plans`` + legacy JSON-планы."""
        require_destructive_db_reset()
        plans_report = self._clear_all_plans()
        return DbResetReport(plans=plans_report)

    def reset_calendar_only(self) -> DbResetReport:
        """Частичное обнуление: только производственный календарь."""
        require_destructive_db_reset()
        calendar_reset = self._reset_calendar()
        return DbResetReport(calendar_reset=calendar_reset)

    # ---------- Recover ----------

    def recover_stuck_plates(self) -> RecoverPlatesResponse:
        """Возвращает плиты из статуса 'в плане' в 'в производстве'."""
        recovered = kp_db_plates.recover_stuck_plates(self.db_path)
        return RecoverPlatesResponse(recovered_records=int(recovered))

    # ---------- Internals ----------

    def _clear_all_plans(self) -> dict[str, int]:
        """Удаляет SQLite-планы и legacy JSON-файлы планов с метаданными.

        Пути берёт из ``Settings``: ``plans_dir``, ``plans_metadata_path``,
        ``current_plan_path``.
        """
        report: dict[str, int] = {
            "sqlite_plans": 0,
            "current_plan": 0,
            "metadata": 0,
            "plan_files": 0,
            "total": 0,
        }

        try:
            report["sqlite_plans"] = self.plan_repository.delete_all_plans()

            current_plan_path: Path = self.settings.current_plan_path
            if current_plan_path.exists():
                current_plan_path.unlink()
                report["current_plan"] = 1

            metadata_path: Path = self.settings.plans_metadata_path
            if metadata_path.exists():
                metadata_path.unlink()
                report["metadata"] = 1

            plans_dir: Path = self.settings.plans_dir
            if plans_dir.exists() and plans_dir.is_dir():
                plan_files = [p for p in plans_dir.glob("*.json")]
                report["plan_files"] = len(plan_files)
                shutil.rmtree(plans_dir, ignore_errors=True)
            plans_dir.mkdir(parents=True, exist_ok=True)
        except Exception:
            logger.exception("Ошибка при очистке JSON-планов производства")
            raise

        report["total"] = (
            report["sqlite_plans"]
            + report["current_plan"]
            + report["metadata"]
            + report["plan_files"]
        )
        return report

    def _reset_calendar(self) -> bool:
        try:
            self.calendar_repository.save_raw(dict(EMPTY_CALENDAR))
            return True
        except Exception:
            logger.exception("Ошибка при сбросе production-календаря")
            raise

    def _count_legacy_json_files(self, *directories: Path) -> int:
        total = 0
        for directory in directories:
            if directory.is_dir():
                total += len(list(directory.glob("*.json")))
        return total

    def _clear_archived_legacy(self) -> dict[str, int]:
        """Best-effort cleanup of bot_archived/data plan artifacts (full reset only)."""
        report: dict[str, int] = {
            "archived_plan_files": 0,
            "archived_metadata": 0,
            "archived_calendar": 0,
        }
        archived_dir = self.settings.archived_data_dir
        archived_plans_dir = archived_dir / "plans"

        try:
            if archived_plans_dir.exists() and archived_plans_dir.is_dir():
                plan_files = list(archived_plans_dir.glob("*.json"))
                report["archived_plan_files"] = len(plan_files)
                shutil.rmtree(archived_plans_dir, ignore_errors=True)
            archived_plans_dir.mkdir(parents=True, exist_ok=True)

            metadata_path = archived_dir / "plans_metadata.json"
            if metadata_path.exists():
                metadata_path.unlink()
                report["archived_metadata"] = 1

            calendar_path = archived_dir / "work_calendar.json"
            if calendar_path.exists():
                calendar_path.unlink()
                report["archived_calendar"] = 1
        except Exception:
            logger.exception("Ошибка при очистке legacy-данных в bot_archived/data")
            raise

        return report


def _normalize_int_dict(raw: dict | None) -> dict[str, int]:
    """Преобразует dict в словарь со строковыми ключами и int-значениями."""
    if not raw:
        return {}
    result: dict[str, int] = {}
    for key, value in raw.items():
        try:
            result[str(key)] = int(value)
        except (TypeError, ValueError):
            continue
    return result
