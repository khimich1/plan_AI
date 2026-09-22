"""Hash раскладки по layout-полям.

Приёмка PR-1/PR-2: hash не должен меняться при добавлении атрибутивных полей
(``concrete_grade``, ``kp_id``, ``kp_plate_id``) — иначе propagation марки
ломает сверку по построению (D-hash).
"""
from __future__ import annotations

import hashlib
import json
from typing import Any

ATTRIBUTION_KEYS = frozenset({"concrete_grade", "kp_id", "kp_plate_id"})


def layout_canonical(sequence: list[Any] | tuple[Any, ...]) -> list[Any]:
    """Каноническая проекция sequence: атрибутивные ключи вырезаны рекурсивно."""
    return [_strip_attribution(item) for item in sequence]


def _strip_attribution(obj: Any) -> Any:
    if isinstance(obj, dict):
        return {
            key: _strip_attribution(value)
            for key, value in obj.items()
            if key not in ATTRIBUTION_KEYS
        }
    if isinstance(obj, (list, tuple)):
        return [_strip_attribution(item) for item in obj]
    return obj


def hash_layout_sequence(sequence: list[Any] | tuple[Any, ...] | None) -> str:
    """SHA-256 канонического JSON sequence без атрибутивных полей."""
    items = list(sequence or [])
    rendered = json.dumps(
        layout_canonical(items),
        sort_keys=True,
        ensure_ascii=False,
        default=str,
        separators=(",", ":"),
    )
    return hashlib.sha256(rendered.encode("utf-8")).hexdigest()


def golden_plan() -> dict[str, Any]:
    """Воспроизводимый вход ``_build_sequence_from_plan`` (solid + split + transverse)."""
    return {
        "plate_assignments": [
            {
                "plate_uid": "plate-solid",
                "kp_id": 1,
                "plate_name": "Плита ПБ 55-12-8п",
                "length": 5.5,
                "width": 1200,
                "load_code": 8,
                "concrete_grade": "М500",
            }
        ],
        "primary_cuts": [
            {
                "qty": 1,
                "width": 1200,
                "rest": 0,
                "lengths": [5.5],
                "plate_uids": ["plate-solid"],
                "load_code": 8,
                "kp_id": 1,
                "plate_name": "Плита ПБ 55-12-8п",
                "concrete_grade": "М500",
            },
            {
                "qty": 1,
                "width": 1200,
                "rest": 200,
                "lengths": [6.0],
                "plate_uids": ["plate-split"],
                "load_code": 8,
                "kp_id": 2,
                "plate_name": "Плита ПБ 60-12-8п",
                "concrete_grade": "М400",
            },
            {
                "qty": 1,
                "width": 1200,
                "rest": 0,
                "lengths": [4.2],
                "plate_uids": ["plate-transverse"],
                "load_code": 8,
                "kp_id": 3,
                "plate_name": "Плита ПБ 42-12-8п",
                "concrete_grade": "М500",
            },
        ],
        "secondary_cuts": [],
        "transverse_cuts": [
            {
                "source_length": 4.2,
                "source_width": 1200,
                "target_length": 3.0,
                "remainder": 1.2,
            }
        ],
    }


def _default_plate_label(length: float, width_m: float, load_code: int | None = None) -> str:
    return f"{length:g}x{int(round(width_m * 1000))}-{load_code or 8}"


def build_golden_sequence() -> list[dict[str, Any]]:
    """Sequence эталонного плана через ``_build_sequence_from_plan``."""
    from viz_modules.layout_sequence.from_plan import _build_sequence_from_plan

    return _build_sequence_from_plan(golden_plan(), _default_plate_label, {})
