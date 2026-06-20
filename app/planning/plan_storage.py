"""
Загрузка, сохранение и метаданные производственных планов.
"""
import ast
import json
import logging
import os
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import List, Optional

logger = logging.getLogger(__name__)

_PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

from core import kp_db
from core.serialization import strip_plate_audit_from_plan

BOT_DIR = _PROJECT_ROOT / "bot"
PLANS_DIR = BOT_DIR / "data" / "plans"
PLANS_METADATA_PATH = BOT_DIR / "data" / "plans_metadata.json"

MAX_TRACKS_PER_DAY = 5

_PLAN_ID_RE = re.compile(r"^[A-Za-z0-9_-]+$")
_MAX_PLAN_ID_LEN = 200


class InvalidPlanIdError(ValueError):
    """Недопустимый plan_id (в т.ч. попытка path traversal)."""


def _validate_plan_id(plan_id: Optional[str]) -> None:
    if not plan_id or not isinstance(plan_id, str):
        raise InvalidPlanIdError("plan_id must be a non-empty string")
    if len(plan_id) > _MAX_PLAN_ID_LEN or not _PLAN_ID_RE.fullmatch(plan_id):
        raise InvalidPlanIdError(f"bad plan_id: {plan_id!r}")


def _plans_dir_resolved() -> Path:
    return PLANS_DIR.resolve()


def _resolved_path_under_plans(candidate: Path, plans_root: Path) -> Path:
    resolved = candidate.resolve()
    try:
        resolved.relative_to(plans_root)
    except ValueError:
        logger.warning("[PLANS] Отклонён путь вне каталога планов: %s", resolved)
        raise InvalidPlanIdError("path escapes plans directory") from None
    return resolved


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
    real = len(day_data.get('tracks', []) or [])
    saved = day_data.get('saved_tracks_count')
    if saved is not None and saved != real:
        logger.warning(
            "[DAY_TRACKS] Рассинхрон: saved_tracks_count=%s, len(tracks)=%s",
            saved, real,
        )
    return real


def ensure_plans_dir():
    """Создаёт папку для планов, если её нет."""
    PLANS_DIR.mkdir(parents=True, exist_ok=True)


def create_plan_id() -> str:
    """Генерирует уникальный ID плана на основе текущего времени."""
    return f"plan_{datetime.now().strftime('%Y%m%d_%H%M%S')}"


def load_plans_metadata() -> dict:
    """Загружает метаданные всех планов."""
    if not PLANS_METADATA_PATH.exists():
        return {"plans": [], "active_plan_id": None}

    try:
        with open(PLANS_METADATA_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        logger.exception(f"Ошибка загрузки метаданных планов: {e}")
        return {"plans": [], "active_plan_id": None}


def save_plans_metadata(metadata: dict):
    """Сохраняет метаданные планов."""
    ensure_plans_dir()
    try:
        with open(PLANS_METADATA_PATH, 'w', encoding='utf-8') as f:
            json.dump(metadata, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logger.exception(f"Ошибка сохранения метаданных планов: {e}")


def get_plan_path(plan_id: str) -> Path:
    """Возвращает абсолютный путь к файлу плана внутри PLANS_DIR."""
    _validate_plan_id(plan_id)
    plans_root = _plans_dir_resolved()
    candidate = plans_root / f"{plan_id}.json"
    return _resolved_path_under_plans(candidate, plans_root)


def load_plan(plan_id: str) -> Optional[dict]:
    """Загружает конкретный план по ID."""
    try:
        plan_path = get_plan_path(plan_id)
    except InvalidPlanIdError:
        logger.warning("Отклонён недопустимый plan_id: %r", plan_id)
        return None
    if not plan_path.exists():
        logger.warning(f"План {plan_id} не найден: {plan_path}")
        return None

    try:
        with open(plan_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        logger.exception(f"Ошибка загрузки плана {plan_id}: {e}")
        return None


def _make_plan_json_serializable(plan_data: dict) -> dict:
    """Строит копию плана без несериализуемых в JSON полей."""
    return strip_plate_audit_from_plan(plan_data)


def save_plan(plan_data: dict) -> bool:
    """Сохраняет план в файл."""
    ensure_plans_dir()
    plan_id = plan_data.get('id')
    if not plan_id:
        logger.error("План не содержит ID!")
        return False

    try:
        plan_path = get_plan_path(plan_id)
    except InvalidPlanIdError:
        logger.error("Недопустимый ID плана при сохранении: %r", plan_id)
        return False
    try:
        to_save = _make_plan_json_serializable(plan_data)
        with open(plan_path, 'w', encoding='utf-8') as f:
            json.dump(to_save, f, ensure_ascii=False, indent=2)
        logger.info(f"План {plan_id} сохранён в {plan_path}")
        return True
    except Exception as e:
        logger.exception(f"Ошибка сохранения плана {plan_id}: {e}")
        return False


def delete_plan(plan_id: str) -> bool:
    """Удаляет план и обновляет метаданные."""
    try:
        plan_path = get_plan_path(plan_id)
    except InvalidPlanIdError:
        return False

    db_path = str(_PROJECT_ROOT / "plita.db")
    returned_count = kp_db.return_plan_plates_to_production(plan_id, db_path)
    if returned_count > 0:
        logger.info(
            f"При удалении плана {plan_id}: возвращено {returned_count} записей плит в производство"
        )

    if plan_path.exists():
        try:
            os.remove(plan_path)
        except Exception as e:
            logger.exception(f"Ошибка удаления файла плана {plan_id}: {e}")
            return False

    metadata = load_plans_metadata()
    metadata['plans'] = [p for p in metadata['plans'] if p['id'] != plan_id]

    if metadata.get('active_plan_id') == plan_id:
        metadata['active_plan_id'] = None
        if metadata['plans']:
            metadata['active_plan_id'] = metadata['plans'][0]['id']

    save_plans_metadata(metadata)
    logger.info(f"План {plan_id} удалён")
    return True


def get_active_plan_id() -> Optional[str]:
    """Возвращает ID активного плана."""
    metadata = load_plans_metadata()
    return metadata.get('active_plan_id')


def set_active_plan(plan_id: str) -> bool:
    """Устанавливает план как активный, только если JSON плана существует на диске."""
    try:
        plan_path = get_plan_path(plan_id)
    except InvalidPlanIdError:
        logger.warning("set_active_plan: недопустимый plan_id %r", plan_id)
        return False
    if not plan_path.is_file():
        logger.warning("set_active_plan: файл плана не найден %s", plan_path)
        return False
    metadata = load_plans_metadata()
    metadata['active_plan_id'] = plan_id
    save_plans_metadata(metadata)
    logger.info(f"План {plan_id} установлен как активный")
    return True


def get_active_plan() -> Optional[dict]:
    """Загружает активный план."""
    active_id = get_active_plan_id()
    if not active_id:
        return None
    return load_plan(active_id)


def get_all_active_plan_ids() -> List[str]:
    """Возвращает список ID всех планов, которые сохранены в metadata."""
    metadata = load_plans_metadata()
    plans = metadata.get('plans', [])
    if isinstance(plans, dict):
        return list(plans.keys())
    return [p.get('id') for p in plans if isinstance(p, dict) and p.get('id')]


def update_plan_metadata(plan: dict):
    """Обновляет запись плана в метаданных."""
    metadata = load_plans_metadata()
    plan_id = plan['id']

    total_days = len(plan.get('days', {}))
    total_tracks = sum(
        count_day_tracks(day)
        for day in plan.get('days', {}).values()
    )

    plan_meta = {
        'id': plan_id,
        'name': plan.get('name', f'План {plan_id}'),
        'created_at': plan.get('created_at', datetime.now().strftime('%Y-%m-%d %H:%M:%S')),
        'start_date': plan.get('start_date', ''),
        'total_days': total_days,
        'tracks_count': plan.get('tracks_count', 5),
        'total_tracks': total_tracks
    }

    found = False
    for i, existing in enumerate(metadata['plans']):
        if existing['id'] == plan_id:
            metadata['plans'][i] = plan_meta
            found = True
            break

    if not found:
        metadata['plans'].append(plan_meta)

    save_plans_metadata(metadata)
