"""Чистое ядро «коммита» производственного плана в БД.

Содержит бизнес-логику, общую для Telegram-бота и веб-сервиса:

1. :func:`count_assigned_plates` — подсчитывает по источникам
   (``primary``/``secondary``/``rescue``) точные идентичности плит
   ``(kp_id, plate_name)``.
2. :func:`distribute_assigned_plates_to_orders` — распределяет полученные
   счётчики по строкам ``orders_2d`` (совпадает с тем, по чему
   ``kp_db.mark_plates_as_planned`` обновляет БД).
3. :func:`commit_plan_plates` — валидирует распределение и помечает плиты
   как ``'в плане'``. При ошибке откатывает изменения.

Модуль не зависит от aiogram и FastAPI: только ``core`` и стандартная
библиотека. Сохранение плана на диск (``plan_manager.save_plan``) остаётся
ответственностью вызывающего слоя (хендлера или сервиса), чтобы UI-часть
могла добавить свои шаги (сообщения пользователю, генерация диаграмм).
"""
from __future__ import annotations

import logging
from collections import Counter
from dataclasses import dataclass, field
from typing import Any

import core.config_and_data as cfg
from core import kp_db

logger = logging.getLogger(__name__)


OrderIdentity = tuple[int | None, str]


class PlanCommitError(RuntimeError):
    """Доменная ошибка при коммите плана в БД.

    Возникает, если optimizer-плиты не удалось сопоставить с заказами,
    если после распределения остались неиспользованные плиты или если
    ``mark_plates_as_planned`` вернул расхождение ожидаемого и фактического
    количества. При такой ошибке изменения в БД откатываются автоматически.
    """


@dataclass
class CommitResult:
    """Статистика успешного коммита плана.

    Поля:
        plates_marked: сколько записей ``kp_plates`` переведено в статус «в плане».
        plates_skipped: сколько позиций пропущено (qty_to_mark == 0, плита потеряна).
        plates_failed: сколько вызовов ``mark_plates_as_planned`` вернули success=False.
        plates_mismatched: сколько вызовов обработали qty меньше, чем ожидалось.
        lost_plates: список позиций ``orders_2d``, закрытых частично.
    """

    plates_marked: int = 0
    plates_skipped: int = 0
    plates_failed: int = 0
    plates_mismatched: int = 0
    lost_plates: list[dict[str, Any]] = field(default_factory=list)


def _make_order_identity(order: dict[str, Any]) -> OrderIdentity:
    """Возвращает ключ позиции в терминах БД (соответствует ``mark_plates_as_planned``)."""
    return order.get("kp_id"), str(order.get("plate_name") or "")


def count_assigned_plates(
    optimization_result: dict[str, Any],
    all_tracks_list: list[dict[str, Any]],
) -> tuple[
    dict[str, dict[OrderIdentity, int]],
    dict[str, list[dict[str, Any]]],
]:
    """Считает плиты по точной идентичности ``(kp_id, plate_name)``.

    Args:
        optimization_result: результат оптимизации, содержит ``plate_assignments``
            — plates, которые оптимизатор сопоставил с конкретной позицией заказа.
        all_tracks_list: итоговый список дорожек. Используется только для
            учёта RESCUE-дорожек, которые добавляются после оптимизации.

    Returns:
        Кортеж из двух словарей:

        - ``assigned_counts_by_source``: ``{source: {(kp_id, plate_name): qty}}``
          для корректно сопоставленных плит по источникам
          ``primary`` / ``secondary`` / ``rescue``.
        - ``unmapped_assignments_by_source``: ``{source: [assignment, ...]}``
          для записей без ``kp_id`` или ``plate_name`` (их нельзя пометить в БД).
    """
    assigned_counts_by_source: dict[str, Counter[OrderIdentity]] = {
        "primary": Counter(),
        "secondary": Counter(),
        "rescue": Counter(),
    }
    unmapped_assignments_by_source: dict[str, list[dict[str, Any]]] = {
        "primary": [],
        "secondary": [],
        "rescue": [],
    }

    for assignment in optimization_result.get("plate_assignments", []) or []:
        source = str(assignment.get("source") or "unknown")
        if source not in assigned_counts_by_source:
            assigned_counts_by_source[source] = Counter()
            unmapped_assignments_by_source[source] = []
        identity = _make_order_identity(assignment)
        if identity[0] and identity[1]:
            assigned_counts_by_source[source][identity] += 1
        else:
            unmapped_assignments_by_source[source].append(
                {
                    "source": assignment.get("source", "unknown"),
                    "length": round(float(assignment.get("length", 0) or 0), 2),
                    "width": assignment.get("width"),
                    "load_code": assignment.get("load_code"),
                    "kp_id": assignment.get("kp_id"),
                    "plate_name": assignment.get("plate_name"),
                    "identity_match_type": assignment.get("identity_match_type"),
                }
            )

    for track in all_tracks_list:
        if track.get("label") != "РЕСКЬЮ":
            continue
        for item in track.get("items", []) or []:
            if not item:
                continue
            identity = _make_order_identity(item)
            if identity[0] and identity[1]:
                assigned_counts_by_source["rescue"][identity] += 1
            else:
                unmapped_assignments_by_source["rescue"].append(
                    {
                        "source": "rescue",
                        "length": round(float(item.get("length", 0) or 0), 2),
                        "width": item.get("width"),
                        "load_code": item.get("load_code"),
                        "kp_id": item.get("kp_id"),
                        "plate_name": item.get("plate_name"),
                        "rescue_order_missing": item.get("rescue_order_missing", False),
                    }
                )

    return (
        {source: dict(counter) for source, counter in assigned_counts_by_source.items()},
        unmapped_assignments_by_source,
    )


def _allocate_counts_to_orders(
    orders_2d: list[dict[str, Any]],
    source_counts: dict[OrderIdentity, int],
    qty_to_mark_by_index: list[int],
) -> tuple[list[int], dict[OrderIdentity, int]]:
    """Добавляет счётчики одного источника в ещё не закрытый спрос."""
    remaining_by_identity: Counter[OrderIdentity] = Counter(source_counts)

    for idx, order in enumerate(orders_2d):
        identity = _make_order_identity(order)
        qty_ordered = int(order.get("qty", 1) or 0)
        qty_missing = max(qty_ordered - qty_to_mark_by_index[idx], 0)
        if qty_missing <= 0:
            continue
        qty_available = remaining_by_identity.get(identity, 0)
        if qty_available <= 0:
            continue
        take = min(qty_missing, qty_available)
        qty_to_mark_by_index[idx] += take
        remaining_by_identity[identity] -= take

    leftovers = {key: qty for key, qty in remaining_by_identity.items() if qty > 0}
    return qty_to_mark_by_index, leftovers


def distribute_assigned_plates_to_orders(
    orders_2d: list[dict[str, Any]],
    assigned_counts_by_source: dict[str, dict[OrderIdentity, int]],
) -> tuple[
    list[dict[str, Any]],
    list[tuple[dict[str, Any], int]],
    dict[str, dict[OrderIdentity, int]],
]:
    """Распределяет уже сопоставленные плиты по строкам ``orders_2d``.

    Возвращает кортеж:

    - ``lost_plates`` — заказы, у которых ``qty_to_mark < qty``.
    - ``orders_with_qty`` — список пар ``(order, qty_to_mark)``
      в исходном порядке ``orders_2d``.
    - ``leftovers_by_source`` — что осталось в каждом источнике после распределения.
    """
    qty_to_mark_by_index: list[int] = [0] * len(orders_2d)
    leftovers_by_source: dict[str, dict[OrderIdentity, int]] = {}

    for source in ("primary", "secondary", "rescue"):
        source_counts = assigned_counts_by_source.get(source) or {}
        qty_to_mark_by_index, leftovers = _allocate_counts_to_orders(
            orders_2d,
            source_counts,
            qty_to_mark_by_index,
        )
        leftovers_by_source[source] = leftovers

    lost: list[dict[str, Any]] = []
    orders_with_qty: list[tuple[dict[str, Any], int]] = []

    for idx, order in enumerate(orders_2d):
        qty_ordered = int(order.get("qty", 1) or 0)
        qty_to_mark = qty_to_mark_by_index[idx]
        orders_with_qty.append((order, qty_to_mark))
        if qty_to_mark < qty_ordered:
            lost.append(
                {
                    "kp_id": order.get("kp_id"),
                    "plate_name": order.get("plate_name", ""),
                    "qty_lost": qty_ordered - qty_to_mark,
                    "load_code": cfg.normalize_load_code(order.get("load_code", 8)),
                }
            )

    return lost, orders_with_qty, leftovers_by_source


def commit_plan_plates(
    *,
    plan_id: str,
    orders_2d: list[dict[str, Any]],
    optimization_result: dict[str, Any],
    all_tracks_list: list[dict[str, Any]],
    db_path: str,
) -> CommitResult:
    """Помечает плиты как «в плане» и валидирует результат.

    Пайплайн:

    1. :func:`count_assigned_plates` → счётчики по источникам.
    2. :func:`distribute_assigned_plates_to_orders` → ``orders_with_qty``.
    3. Валидация: ``optimizer_unmapped`` и ``optimizer_extra`` должны быть
       пустыми (иначе нельзя гарантировать корректность пометки).
    4. Цикл ``kp_db.mark_plates_as_planned`` по каждой позиции заказа.
    5. При любой проблеме — ``kp_db.return_plan_plates_to_production`` и
       выброс :class:`PlanCommitError`.

    Args:
        plan_id: идентификатор плана (используется при записи ``plan_id`` в БД
            и при откате).
        orders_2d: список заказов в виде ``{kp_id, plate_name, qty, ...}``.
        optimization_result: результат оптимизации с ``plate_assignments``.
        all_tracks_list: итоговые дорожки (нужны для учёта RESCUE-плит).
        db_path: путь к ``plita.db``.

    Returns:
        :class:`CommitResult` со статистикой пометки.

    Raises:
        PlanCommitError: если валидация не прошла или фактическая пометка
            не соответствует ожидаемой.
    """
    assigned_counts_by_source, unmapped_assignments_by_source = count_assigned_plates(
        optimization_result=optimization_result,
        all_tracks_list=all_tracks_list,
    )

    lost_plates, orders_with_qty, leftovers_by_source = distribute_assigned_plates_to_orders(
        orders_2d=orders_2d,
        assigned_counts_by_source=assigned_counts_by_source,
    )

    source_totals = {
        source: sum(counts.values())
        for source, counts in assigned_counts_by_source.items()
        if counts
    }
    if source_totals:
        logger.info("[PLAN_COMMIT] Источники exact identity перед пометкой: %s", source_totals)

    if lost_plates:
        lost_info = ", ".join(
            f"{lp['plate_name']} x{lp['qty_lost']}" for lp in lost_plates[:3]
        )
        if len(lost_plates) > 3:
            lost_info += f" и ещё {len(lost_plates) - 3}..."
        logger.warning(
            "[PLAN_COMMIT] Обнаружены потерянные плиты (НЕ будут помечены): %s",
            lost_info,
        )

    optimizer_unmapped = (
        (unmapped_assignments_by_source.get("primary") or [])
        + (unmapped_assignments_by_source.get("secondary") or [])
    )
    rescue_unmapped = unmapped_assignments_by_source.get("rescue") or []

    if optimizer_unmapped:
        logger.error(
            "[PLAN_COMMIT] Найдены optimizer-плиты без точной привязки к заказу: %s",
            optimizer_unmapped[:10],
        )
        raise PlanCommitError(
            "Не удалось сопоставить часть плит с конкретными заказами."
        )

    if rescue_unmapped:
        logger.warning(
            "[PLAN_COMMIT] RESCUE-плиты без exact identity не будут участвовать в пометке БД: %s",
            rescue_unmapped[:10],
        )

    optimizer_leftovers = {
        "primary": leftovers_by_source.get("primary") or {},
        "secondary": leftovers_by_source.get("secondary") or {},
    }
    rescue_leftovers = leftovers_by_source.get("rescue") or {}
    optimizer_extra = {
        source: leftovers
        for source, leftovers in optimizer_leftovers.items()
        if leftovers
    }

    if optimizer_extra:
        logger.error(
            "[PLAN_COMMIT] После распределения остались optimizer-плиты без заказа: %s",
            optimizer_extra,
        )
        raise PlanCommitError(
            "В плане обнаружены плиты, не сопоставленные с заказами."
        )

    if rescue_leftovers:
        logger.warning(
            "[PLAN_COMMIT] RESCUE-плиты сверх уже покрытого спроса игнорируются при пометке БД: %s",
            rescue_leftovers,
        )

    result = CommitResult(lost_plates=lost_plates)
    marked_any = False

    for order, qty_to_mark in orders_with_qty:
        kp_id = order.get("kp_id")
        plate_name = order.get("plate_name")
        qty_ordered = int(order.get("qty", 1) or 0)

        if qty_to_mark <= 0:
            if qty_ordered > 0:
                result.plates_skipped += 1
                logger.info(
                    "[PLAN_COMMIT] Пропускаем потерянную плиту: КП #%s, %s x%s",
                    kp_id,
                    plate_name,
                    qty_ordered,
                )
            continue

        if qty_to_mark < qty_ordered:
            logger.info(
                "[PLAN_COMMIT] Частичная потеря: КП #%s, %s — помечаем %s из %s",
                kp_id,
                plate_name,
                qty_to_mark,
                qty_ordered,
            )

        if not (kp_id and plate_name and qty_to_mark > 0):
            continue

        mark_result = kp_db.mark_plates_as_planned(
            kp_id=kp_id,
            plate_name=plate_name,
            qty_to_plan=qty_to_mark,
            plan_id=plan_id,
            db_path=db_path,
        )
        if mark_result.get("success"):
            processed_count = int(mark_result.get("processed_count", 0) or 0)
            result.plates_marked += processed_count
            marked_any = marked_any or processed_count > 0
            if processed_count != qty_to_mark:
                result.plates_mismatched += 1
                logger.error(
                    "[PLAN_COMMIT] Расхождение при пометке плиты: КП #%s, %s. "
                    "Ожидалось %s, фактически помечено %s.",
                    kp_id,
                    plate_name,
                    qty_to_mark,
                    processed_count,
                )
        else:
            result.plates_failed += 1
            logger.error(
                "[PLAN_COMMIT] mark_plates_as_planned вернул ошибку: КП #%s, %s. Детали: %s",
                kp_id,
                plate_name,
                mark_result,
            )

    logger.info(
        "[PLAN_COMMIT] План %s: помечено %s плит, пропущено %s",
        plan_id,
        result.plates_marked,
        result.plates_skipped,
    )

    if result.plates_failed > 0 or result.plates_mismatched > 0:
        logger.error(
            "[PLAN_COMMIT] Ошибки при пометке плит. failed=%s, mismatched=%s. Откатываю...",
            result.plates_failed,
            result.plates_mismatched,
        )
        try:
            kp_db.return_plan_plates_to_production(plan_id, db_path)
        except Exception:
            logger.exception(
                "[PLAN_COMMIT] Ошибка при откате плит для плана %s",
                plan_id,
            )
        raise PlanCommitError(
            f"Не удалось корректно пометить плиты в БД: "
            f"failed={result.plates_failed}, mismatched={result.plates_mismatched}."
        )

    return result


__all__ = [
    "CommitResult",
    "PlanCommitError",
    "commit_plan_plates",
    "count_assigned_plates",
    "distribute_assigned_plates_to_orders",
]
