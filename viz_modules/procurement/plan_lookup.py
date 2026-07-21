from __future__ import annotations

from collections.abc import Mapping

from .constants import WIDE_EPS, WIDE_WIDTH_M
from .plan_snapshot import parse_cascading_plan, snapshot_to_trim_dict


def _is_wide_width(width_m: float, *, threshold_m: float = WIDE_WIDTH_M, eps: float = WIDE_EPS) -> bool:
    """True для плит шире 12 дм (> 1.2 м)."""
    return float(width_m) > (threshold_m + eps)


def _block_ab_key(width_m: float) -> int:
    """Ключ двухблочной схемы: 0=Блок A (обычные), 1=Блок B (широкие)."""
    return 1 if _is_wide_width(width_m) else 0

def _length_dm_raw_from_m(length_m: float) -> str:
    """Вычисляет строку длины в дм (как в марке) из длины в метрах.

    Зеркало логики make_plate_name при отсутствии length_dm_raw:
    5.71 → '57,1', 5.7 → '57', 6.88 → '68,8'.
    """
    length_dm_val = length_m * 10
    if abs(length_dm_val - round(length_dm_val)) < 0.01:
        return str(int(round(length_dm_val)))
    return f'{length_dm_val:.1f}'.rstrip('0').rstrip('.').replace('.', ',')


def _is_same_length(lengths: list, target_len: float, tolerance: float = 0.05) -> bool:
    """Проверяет, подходит ли операция под длину плиты."""
    if not lengths:
        return True
    return any(abs(float(v) - target_len) < tolerance for v in lengths)


def _find_plan_for_plate(
    load_code: int,
    length: float,
    width_mm: int,
    name: str,
    debug_tag: str,
    *,
    plan_by_load: Mapping | None = None,
):
    """Ищет план оптимизации для конкретной плиты по нагрузке/длине/ширине."""
    from core.optimization import LOAD_TO_REINFORCEMENT_MAP, OPT_CASCADING_PLAN_BY_LOAD
    import math

    current_plan = None
    load_key = int(math.floor(load_code)) if isinstance(load_code, (int, float)) else 8

    src_plan_by_load = plan_by_load if plan_by_load is not None else OPT_CASCADING_PLAN_BY_LOAD

    if src_plan_by_load and LOAD_TO_REINFORCEMENT_MAP and load_key in LOAD_TO_REINFORCEMENT_MAP:
        for reinforcement_key in LOAD_TO_REINFORCEMENT_MAP[load_key]:
            plan_raw = src_plan_by_load.get(reinforcement_key)
            if not plan_raw:
                continue

            snap = parse_cascading_plan(plan_raw if isinstance(plan_raw, Mapping) else {})
            for ord_row in snap.orders_requested:
                try:
                    o_len = float(ord_row.length)
                    o_width = int(float(ord_row.width))
                    o_load = ord_row.load_code if ord_row.load_code is not None else load_key
                except (TypeError, ValueError):
                    continue

                if (
                    abs(o_len - length) < 0.05 and
                    o_width == width_mm and
                    int(math.floor(float(o_load))) == load_key
                ):
                    current_plan = snapshot_to_trim_dict(snap)
                    print(
                        f'[DEBUG] {debug_tag}: нашёл план для {name} — '
                        f'нагрузка {load_key}п, армирование {reinforcement_key}'
                    )
                    break
            if current_plan:
                break
    return current_plan, load_key

