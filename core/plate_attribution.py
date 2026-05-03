"""Атрибуция плит к заказам после оптимизации.

Этот модуль содержит вспомогательные функции, которые делают
``optimization_result["plate_assignments"]`` и ``track["items"]``
исчерпывающими списками плит с заполненной identity ``(kp_id, plate_name)``.

Проблема, которую решает модуль:
    ``core.optimization._next_slot_info`` строит атрибуцию по
    proportional slots от ``demand_2d``. Если оптимизатор произвёл больше
    плит ключа ``(L, W, LC)``, чем рассчитанный спрос (например, из-за
    округлений или edge-кейсов CP-SAT), слоты исчерпываются и оптимизатор
    оставляет ``kp_id=None`` / ``plate_name=None``. Такие записи попадают
    в ``unmapped_assignments_by_source`` и приводят к
    :class:`PlanCommitError`.

Аналогичная история для ``track["items"]`` и вложенных ``secondary_cuts``:
    ``viz_modules.layout_sequence.build_layout_sequence`` копирует identity
    из первичных cut'ов в root-items, но secondary cuts из ``chosen_variant``
    создаются БЕЗ ``kp_id``/``plate_name`` (только с ``target_order_key``).
    Без backfill эти items невидимы для ``commit_plan_plates``
    (``_count_track_items_by_day`` не строит per-day для них), плиты
    помечаются ``status='в плане'`` без ``day_number`` и зависают в БД.

После backfill каждая запись ``plate_assignments`` и каждый item трека
имеет identity, поэтому:
- ``count_assigned_plates`` корректно учитывает все плиты по identity;
- ``_count_track_items_by_day`` распределяет каждую плиту по дням;
- ``kp_plate_id`` записывается в каждый item;
- DB-путь ``build_day_view_detail`` работает для 100% треков.
"""
from __future__ import annotations

import logging
from collections import defaultdict
from typing import Any

from core.config_and_data import canonical_plate_key, normalize_load_code

logger = logging.getLogger(__name__)


ConsumedMap = dict[tuple, dict[tuple[int, str], int]]


def _new_consumed() -> ConsumedMap:
    return defaultdict(lambda: defaultdict(int))


def _build_orders_index(
    orders_2d: list[dict[str, Any]] | None,
) -> dict[tuple, list[dict[str, Any]]]:
    """Группирует ``orders_2d`` по ``canonical_plate_key``."""
    index: dict[tuple, list[dict[str, Any]]] = defaultdict(list)
    for order in orders_2d or []:
        if order.get("kp_id") is None:
            continue
        key = canonical_plate_key(
            order.get("length", 0),
            order.get("width", 1200),
            order.get("load_code", 8),
        )
        index[key].append(order)
    return index


def _pick_best_order(
    candidates: list[dict[str, Any]],
    consumed_by_pair: dict[tuple[int, str], int],
) -> dict[str, Any] | None:
    """Выбирает заказ с наибольшим непокрытым спросом ``qty - consumed``."""
    best_order: dict[str, Any] | None = None
    best_remaining = -1
    for order in candidates:
        kp_id = order.get("kp_id")
        plate_name = order.get("plate_name") or ""
        if kp_id is None:
            continue
        qty_ordered = int(order.get("qty", 1) or 0)
        already = consumed_by_pair.get((int(kp_id), str(plate_name)), 0)
        remaining = qty_ordered - already
        if remaining > best_remaining:
            best_remaining = remaining
            best_order = order
    if best_order is None and candidates:
        best_order = candidates[0]
    return best_order


def backfill_assignment_identity(
    plate_assignments: list[dict[str, Any]],
    orders_2d: list[dict[str, Any]],
    *,
    consumed: ConsumedMap | None = None,
) -> int:
    """Заполняет ``kp_id``/``plate_name`` у записей ``plate_assignments``,
    у которых их нет, по заказам из ``orders_2d``.

    Записи матчатся по каноническому ключу
    ``canonical_plate_key(length, width, load_code)``. Для каждой записи без
    identity выбирается заказ из ``orders_2d`` с тем же ключом, у которого
    остался наибольший непокрытый спрос (``qty - уже_атрибутировано``).

    Args:
        plate_assignments: список плит из ``optimization_result``.
            Изменяется на месте.
        orders_2d: список заказов в формате ``{kp_id, plate_name, length,
            width (mm), load_code, qty, ...}``.
        consumed: опциональный counter ``{canonical_key:
            {(kp_id, plate_name): int}}``. По умолчанию ``None`` —
            создаётся локальный counter. Параметр оставлен для тестов
            и продвинутых сценариев. Внимание: этот counter НЕ нужно
            шарить с ``backfill_track_items_identity`` — assignments и
            track-items это два независимых взгляда на ОДНИ и ТЕ ЖЕ
            плиты, и каждый ведёт свой подсчёт атрибуции.

    Returns:
        Количество backfill-записей (``identity_match_type='backfilled'``).
    """
    if not plate_assignments:
        return 0

    orders_by_key = _build_orders_index(orders_2d)
    if not orders_by_key:
        return 0

    if consumed is None:
        consumed = _new_consumed()

    for assignment in plate_assignments:
        kp_id = assignment.get("kp_id")
        plate_name = assignment.get("plate_name")
        if kp_id and plate_name:
            key = canonical_plate_key(
                assignment.get("length", 0),
                assignment.get("width", 1200),
                assignment.get("load_code", 8),
            )
            consumed[key][(int(kp_id), str(plate_name))] += 1

    backfilled = 0
    for assignment in plate_assignments:
        if assignment.get("kp_id") and assignment.get("plate_name"):
            continue

        key = canonical_plate_key(
            assignment.get("length", 0),
            assignment.get("width", 1200),
            assignment.get("load_code", 8),
        )
        candidates = orders_by_key.get(key) or []
        if not candidates:
            logger.debug(
                "[BACKFILL] Не нашли заказ под ключ %s для assignment %s",
                key, assignment,
            )
            continue

        best_order = _pick_best_order(candidates, consumed[key])
        if best_order is None:
            continue

        assignment["kp_id"] = best_order.get("kp_id")
        assignment["plate_name"] = best_order.get("plate_name") or ""
        assignment["identity_match_type"] = "backfilled"
        consumed[key][(int(assignment["kp_id"]), str(assignment["plate_name"]))] += 1
        backfilled += 1

    if backfilled:
        logger.info(
            "[BACKFILL] Восстановлена identity для %s assignment-записей "
            "(slot_exhausted/secondary_unmapped)",
            backfilled,
        )

    return backfilled


def _root_item_key(item: dict[str, Any]) -> tuple | None:
    """Канонический ключ для root-item трека (mode in {solid, split, transverse}).

    Для ``split`` берём ``main_w`` (ширина основного куска), для ``solid``
    и ``transverse`` — ``width``. Длина — ``item.length``. Возвращает None,
    если не получается построить ключ.
    """
    length = item.get("length")
    if not length:
        return None
    mode = item.get("mode") or "solid"
    if mode == "split":
        width = item.get("main_w")
    else:
        width = item.get("width")
    if width is None:
        return None
    if isinstance(width, (int, float)) and width < 20:
        width_mm = int(round(float(width) * 1000))
    else:
        width_mm = int(round(float(width)))
    load_code = item.get("load_code", 8)
    try:
        return canonical_plate_key(length, width_mm, load_code)
    except Exception:  # noqa: BLE001
        return None


def _secondary_keys_priority(
    sec: dict[str, Any],
    parent_length: float | None,
) -> list[tuple]:
    """Возвращает список candidate-ключей для secondary cut по приоритету.

    Порядок:
    1. ``target_order_key`` (из оптимизатора) — самый точный сигнал.
    2. ``target_length`` + ``width`` + ``load_code`` (для transverse-резов).
    3. ``parent_length`` + ``width`` + ``load_code`` (для narrowing — длина
       совпадает с родителем).
    """
    keys: list[tuple] = []

    tok = sec.get("target_order_key")
    if (
        isinstance(tok, (tuple, list))
        and len(tok) >= 3
        and tok[0] is not None
        and tok[1] is not None
    ):
        try:
            keys.append(canonical_plate_key(tok[0], tok[1], tok[2]))
        except Exception:  # noqa: BLE001
            pass

    width_raw = sec.get("width")
    if width_raw is None:
        return keys
    if isinstance(width_raw, (int, float)) and width_raw < 20:
        width_mm = int(round(float(width_raw) * 1000))
    else:
        width_mm = int(round(float(width_raw)))

    load_code_raw = sec.get("load_code", 8)
    try:
        load_code = normalize_load_code(load_code_raw, default=8)
    except Exception:  # noqa: BLE001
        load_code = 8

    target_length = sec.get("target_length")
    if target_length:
        try:
            keys.append(canonical_plate_key(target_length, width_mm, load_code))
        except Exception:  # noqa: BLE001
            pass

    if parent_length:
        try:
            parent_key = canonical_plate_key(parent_length, width_mm, load_code)
        except Exception:  # noqa: BLE001
            parent_key = None
        if parent_key is not None and parent_key not in keys:
            keys.append(parent_key)

    return keys


def _ensure_load_code_on_secondary(sec: dict[str, Any], parent_load_code: Any) -> int:
    raw = sec.get("load_code")
    if raw is None:
        raw = parent_load_code if parent_load_code is not None else 8
    try:
        return normalize_load_code(raw, default=8)
    except Exception:  # noqa: BLE001
        return 8


def _attribute_one_item(
    item: dict[str, Any],
    candidate_keys: list[tuple],
    orders_by_key: dict[tuple, list[dict[str, Any]]],
    consumed: ConsumedMap,
) -> bool:
    """Пытается найти и проставить identity для одного item/secondary.

    Returns:
        True если identity была успешно проставлена.
    """
    for key in candidate_keys:
        candidates = orders_by_key.get(key) or []
        if not candidates:
            continue
        best_order = _pick_best_order(candidates, consumed[key])
        if best_order is None:
            continue
        kp_id = best_order.get("kp_id")
        plate_name = best_order.get("plate_name") or ""
        if kp_id is None:
            continue
        item["kp_id"] = kp_id
        item["plate_name"] = plate_name
        item["identity_match_type"] = "backfilled"
        consumed[key][(int(kp_id), str(plate_name))] += 1
        return True
    return False


def backfill_track_items_identity(
    tracks_list: list[dict[str, Any]],
    orders_2d: list[dict[str, Any]],
    *,
    consumed: ConsumedMap | None = None,
) -> int:
    """Заполняет ``kp_id``/``plate_name`` у root-items треков и у вложенных
    ``secondary_cuts``, у которых identity отсутствует.

    Зеркало :func:`backfill_assignment_identity`, но работает по структуре
    ``track["items"][*]`` и ``items[*]["secondary_cuts"][*]``.

    Алгоритм для root-item:
        1. Если ``kp_id``+``plate_name`` уже выставлены — пропускаем.
        2. Строим канонический ключ из ``item.length`` и ширины
           (``main_w`` для split, ``width`` для solid/transverse).
        3. Ищем заказ из ``orders_2d`` под этот ключ, выбираем с наибольшим
           непокрытым остатком.
        4. Проставляем identity, увеличиваем consumed-counter.

    Алгоритм для secondary cut:
        1. Если identity уже есть — пропускаем.
        2. Строим список candidate-ключей по приоритету
           (``target_order_key`` → ``target_length`` → ``parent.length``).
        3. Перебираем ключи; первый, для которого нашёлся заказ, определяет
           identity.

    Args:
        tracks_list: список треков. Каждый трек — словарь с полем ``items``.
            Изменяется на месте.
        orders_2d: тот же список заказов, что и для assignments.
        consumed: опциональный counter (см.
            :func:`backfill_assignment_identity`). По умолчанию ``None`` —
            создаётся локальный. Шарить с backfill_assignment_identity
            НЕ нужно, см. примечание там.

    Returns:
        Количество backfill-записей (root + secondary).
    """
    if not tracks_list:
        return 0

    orders_by_key = _build_orders_index(orders_2d)
    if not orders_by_key:
        return 0

    if consumed is None:
        consumed = _new_consumed()

    backfilled = 0

    for track in tracks_list:
        if not isinstance(track, dict):
            continue

        if track.get("label") == "РЕСКЬЮ":
            for item in track.get("items") or []:
                if not isinstance(item, dict):
                    continue
                if item.get("kp_id") and item.get("plate_name"):
                    key = _root_item_key(item)
                    if key is not None:
                        consumed[key][
                            (int(item["kp_id"]), str(item["plate_name"]))
                        ] += 1
            continue

        for item in track.get("items") or []:
            if not isinstance(item, dict):
                continue

            root_key = _root_item_key(item)

            if item.get("kp_id") and item.get("plate_name"):
                if root_key is not None:
                    consumed[root_key][
                        (int(item["kp_id"]), str(item["plate_name"]))
                    ] += 1
            elif root_key is not None:
                if _attribute_one_item(item, [root_key], orders_by_key, consumed):
                    backfilled += 1
                else:
                    logger.debug(
                        "[BACKFILL_ITEMS] Не нашли заказ под root-key %s для item %s",
                        root_key,
                        {
                            "mode": item.get("mode"),
                            "length": item.get("length"),
                            "load_code": item.get("load_code"),
                        },
                    )

            parent_length = item.get("length")
            for sec in item.get("secondary_cuts") or []:
                if not isinstance(sec, dict):
                    continue

                if "load_code" not in sec or sec.get("load_code") is None:
                    sec["load_code"] = _ensure_load_code_on_secondary(
                        sec, item.get("load_code")
                    )

                if sec.get("kp_id") and sec.get("plate_name"):
                    sec_keys = _secondary_keys_priority(sec, parent_length)
                    if sec_keys:
                        consumed[sec_keys[0]][
                            (int(sec["kp_id"]), str(sec["plate_name"]))
                        ] += 1
                    continue

                candidate_keys = _secondary_keys_priority(sec, parent_length)
                if not candidate_keys:
                    logger.debug(
                        "[BACKFILL_ITEMS] secondary без candidate-ключей: %s",
                        {
                            "label": sec.get("label"),
                            "width": sec.get("width"),
                            "target_length": sec.get("target_length"),
                            "target_order_key": sec.get("target_order_key"),
                        },
                    )
                    continue

                if _attribute_one_item(
                    sec, candidate_keys, orders_by_key, consumed
                ):
                    backfilled += 1

    if backfilled:
        logger.info(
            "[BACKFILL_ITEMS] Восстановлена identity у %s track-items "
            "(root + secondary_cuts)",
            backfilled,
        )

    return backfilled


__all__ = [
    "backfill_assignment_identity",
    "backfill_track_items_identity",
    "ConsumedMap",
]
