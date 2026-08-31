"""
Загрузка, сохранение и метаданные производственных планов.

Единый persistence path — SQLite через :class:`app.repositories.plan_repository.PlanRepository`.
JSON-файлы в ``data/plans/`` больше не используются для чтения/записи (см. WP1 / A1).
"""
from __future__ import annotations

import ast
import logging
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import List, Optional

logger = logging.getLogger(__name__)

_PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

from app.core.settings import get_settings
from core import kp_db
from core.serialization import strip_plate_audit_from_plan

BOT_DIR = _PROJECT_ROOT / "bot"
PLANS_DIR = BOT_DIR / "data" / "plans"
PLANS_METADATA_PATH = BOT_DIR / "data" / "plans_metadata.json"

MAX_TRACKS_PER_DAY = 5

_PLAN_ID_RE = re.compile(r"^[A-Za-z0-9_-]+$")
_MAX_PLAN_ID_LEN = 200

_repo_override: object | None = None


class InvalidPlanIdError(ValueError):
    """Недопустимый plan_id (в т.ч. попытка path traversal)."""


def _validate_plan_id(plan_id: Optional[str]) -> None:
    if not plan_id or not isinstance(plan_id, str):
        raise InvalidPlanIdError("plan_id must be a non-empty string")
    if len(plan_id) > _MAX_PLAN_ID_LEN or not _PLAN_ID_RE.fullmatch(plan_id):
        raise InvalidPlanIdError(f"bad plan_id: {plan_id!r}")


def _get_repository():
    from app.repositories.plan_repository import PlanRepository

    if _repo_override is not None:
        return _repo_override
    settings = get_settings()
    return PlanRepository(str(settings.plita_db_path))


def get_repository():
    """Возвращает активный репозиторий планов (для адаптеров и тестов)."""
    return _get_repository()


def convert_lookup_keys_to_tuples(lookup_dict: dict) -> dict:
    """
    Конвертирует строковые ключи lookup-словарей обратно в кортежи или числа.

    JSON не поддерживает кортежи как ключи, поэтому при сохранении
    они конвертируются в строки "(длина, ширина)" или "длина".
    Эта функция восстанавливает исходный формат.
    """
    result = {}

    for key, value in lookup_dict.items():
        if isinstance(key, str):
            try:
                parsed = ast.literal_eval(key)
                if isinstance(parsed, (list, tuple)):
                    result[tuple(parsed)] = value
                else:
                    result[parsed] = value
            except (ValueError, SyntaxError, TypeError):
                result[key] = value
        else:
            result[key] = value

    return result


def count_day_tracks(day_data: dict) -> int:
    """Единая точка truth: сколько дорожек реально в дне."""
    real = len(day_data.get("tracks", []) or [])
    saved = day_data.get("saved_tracks_count")
    if saved is not None and saved != real:
        logger.warning(
            "[DAY_TRACKS] Рассинхрон: saved_tracks_count=%s, len(tracks)=%s",
            saved,
            real,
        )
    return real


def ensure_plans_dir() -> None:
    """Legacy no-op: планы хранятся в SQLite, не в каталоге на диске."""


def create_plan_id() -> str:
    """Генерирует уникальный ID плана на основе текущего времени."""
    return f"plan_{datetime.now().strftime('%Y%m%d_%H%M%S')}"


def load_plans_metadata() -> dict:
    """Загружает метаданные всех планов из SQLite."""
    return _get_repository().list_metadata()


def save_plans_metadata(metadata: dict) -> None:
    """Legacy no-op: метаданные выводятся из ``production_plans`` в SQLite."""
    logger.debug(
        "[PLANS] save_plans_metadata ignored (%d entries); SQLite is authoritative",
        len(metadata.get("plans", []) or []),
    )


def get_plan_path(plan_id: str) -> Path:
    """Legacy path helper (для обратной совместимости импортов)."""
    _validate_plan_id(plan_id)
    return PLANS_DIR / f"{plan_id}.json"


def load_plan(plan_id: str) -> Optional[dict]:
    """Загружает конкретный план по ID из SQLite."""
    try:
        _validate_plan_id(plan_id)
    except InvalidPlanIdError:
        logger.warning("Отклонён недопустимый plan_id: %r", plan_id)
        return None

    try:
        return _get_repository().load_plan(plan_id)
    except Exception as exc:
        logger.exception("Ошибка загрузки плана %s: %s", plan_id, exc)
        return None


def save_plan(plan_data: dict) -> bool:
    """Сохраняет план в SQLite."""
    plan_id = plan_data.get("id")
    if not plan_id:
        logger.error("План не содержит ID!")
        return False

    try:
        _validate_plan_id(plan_id)
        _get_repository().save_plan(plan_data)
        logger.info("План %s сохранён в SQLite", plan_id)
        return True
    except InvalidPlanIdError:
        logger.error("Недопустимый ID плана при сохранении: %r", plan_id)
        return False
    except Exception as exc:
        logger.exception("Ошибка сохранения плана %s: %s", plan_id, exc)
        return False


def delete_plan(plan_id: str) -> bool:
    """Удаляет план из SQLite и возвращает плиты в производство."""
    try:
        _validate_plan_id(plan_id)
    except InvalidPlanIdError:
        return False

    repo = _get_repository()
    if repo.load_plan(plan_id) is None:
        return False

    db_path = str(get_settings().plita_db_path)
    returned_count = kp_db.return_plan_plates_to_production(plan_id, db_path)
    if returned_count > 0:
        logger.info(
            "При удалении плана %s: возвращено %s записей плит в производство",
            plan_id,
            returned_count,
        )

    # SGP-503: плиты на СГП остаются; только снимаем plan_id
    try:
        from app.services.sgp_service import SgpService

        cleared = SgpService(db_path=db_path).clear_plan_links(plan_id)
        if cleared:
            logger.info(
                "При удалении плана %s: обнулён plan_id у %s строк СГП",
                plan_id,
                cleared,
            )
    except Exception:
        logger.exception("Не удалось очистить plan_id СГП для плана %s", plan_id)

    active_id = repo.get_active_plan_id()
    if not repo.delete(plan_id):
        return False

    if active_id == plan_id:
        remaining = [entry["id"] for entry in repo.list_metadata().get("plans", [])]
        if remaining:
            repo.set_active(remaining[0])
        logger.info("Активный план %s удалён; новый active=%s", plan_id, remaining[:1])

    logger.info("План %s удалён из SQLite", plan_id)
    return True


def get_active_plan_id() -> Optional[str]:
    """Возвращает ID активного плана."""
    return _get_repository().get_active_plan_id()


def set_active_plan(plan_id: str) -> bool:
    """Устанавливает план как активный, если он существует в SQLite."""
    try:
        _validate_plan_id(plan_id)
    except InvalidPlanIdError:
        logger.warning("set_active_plan: недопустимый plan_id %r", plan_id)
        return False

    repo = _get_repository()
    if repo.load_plan(plan_id) is None:
        logger.warning("set_active_plan: план %r не найден в SQLite", plan_id)
        return False

    if repo.set_active_plan(plan_id):
        logger.info("План %s установлен как активный", plan_id)
        return True
    return False


def get_active_plan() -> Optional[dict]:
    """Загружает активный план."""
    record = _get_repository().get_active()
    return record["payload"] if record else None


def get_all_active_plan_ids() -> List[str]:
    """Возвращает список ID всех планов из метаданных."""
    metadata = load_plans_metadata()
    plans = metadata.get("plans", [])
    if isinstance(plans, dict):
        return list(plans.keys())
    return [p.get("id") for p in plans if isinstance(p, dict) and p.get("id")]


def update_plan_metadata(plan: dict) -> None:
    """Legacy no-op: метаданные обновляются при сохранении payload в SQLite."""
    logger.debug(
        "[PLANS] update_plan_metadata ignored for plan %s; SQLite is authoritative",
        plan.get("id"),
    )


def _make_plan_json_serializable(plan_data: dict) -> dict:
    """Строит копию плана без несериализуемых в JSON полей."""
    return strip_plate_audit_from_plan(plan_data)
