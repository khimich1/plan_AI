"""Чистое ядро «коммита» производственного плана в БД.

Содержит бизнес-логику, общую для Telegram-бота и веб-сервиса:

1. :func:`count_assigned_plates` — подсчитывает по источникам
   (``primary``/``secondary``/``rescue``) точные идентичности плит
   ``(kp_id, plate_name)``.
2. :func:`distribute_assigned_plates_to_orders` — распределяет полученные
   счётчики по строкам ``orders_2d`` (совпадает с тем, по чему
   ``kp_db_plates.mark_plates_as_planned`` обновляет БД).
3. :func:`commit_plan_plates` — валидирует распределение и помечает плиты
   как ``'в плане'``. При ошибке откатывает изменения.

Модуль не зависит от aiogram и FastAPI: только ``core`` и стандартная
библиотека. Сохранение плана на диск (``app.planning.plan_manager.save_plan``) остаётся
ответственностью вызывающего слоя (хендлера или сервиса), чтобы UI-часть
могла добавить свои шаги (сообщения пользователю, генерация диаграмм).
"""
from __future__ import annotations

import logging
from collections import Counter, defaultdict
from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import date
from typing import Any, Iterable

from core import kp_db_plates, plate_name as _plate_name
from core.domain.plate_order import normalize_load_code
from core.kp_db_common import _connect

logger = logging.getLogger(__name__)


OrderIdentity = tuple[int | None, str]
PromiseSettleFn = Callable[..., Any]


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
    all_tracks_list: list[dict[str, Any]] | None = None,
) -> tuple[
    dict[str, dict[OrderIdentity, int]],
    dict[str, list[dict[str, Any]]],
]:
    """Считает плиты по точной идентичности ``(kp_id, plate_name)``.

    Phase 3 P8: единственный источник учёта — ``plate_assignments``. RESCUE-плиты
    после Phase 2 живут в нём с ``source='rescue'`` (см.
    :func:`core.rescue_tracks.build_rescue_tracks`), поэтому отдельный
    проход по ``all_tracks_list`` больше не нужен и может приводить к
    двойному счёту.

    Args:
        optimization_result: результат оптимизации, содержит ``plate_assignments``
            — flat-список плит с ``source`` ∈ ``primary``/``secondary``/``rescue``.
        all_tracks_list: оставлен для обратной совместимости сигнатуры
            (используется выше по стеку для распределения kp_plate_id по дням).
            Для счёта identity не используется.

    Returns:
        Кортеж из двух словарей:

        - ``assigned_counts_by_source``: ``{source: {(kp_id, plate_name): qty}}``
          для корректно сопоставленных плит по источникам.
        - ``unmapped_assignments_by_source``: ``{source: [assignment, ...]}``
          для записей без ``kp_id`` или ``plate_name`` (их нельзя пометить в БД).
    """
    del all_tracks_list  # после Phase 3 учёт идёт только из plate_assignments
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
                    "rescue_order_missing": assignment.get("rescue_order_missing", False),
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
                    "load_code": normalize_load_code(order.get("load_code", 8)),
                }
            )

    return lost, orders_with_qty, leftovers_by_source


def _accumulate_mark_result(
    mark_result: dict[str, Any],
    result: "CommitResult",
    *,
    kp_id: int,
    plate_name: str,
    expected: int,
) -> None:
    """Аккумулирует статистику от ``mark_plates_as_planned`` в CommitResult."""
    if mark_result.get("success"):
        processed_count = int(mark_result.get("processed_count", 0) or 0)
        result.plates_marked += processed_count
        if processed_count != expected:
            result.plates_mismatched += 1
            logger.error(
                "[PLAN_COMMIT] Расхождение при пометке плиты: КП #%s, %s. "
                "Ожидалось %s, фактически помечено %s.",
                kp_id,
                plate_name,
                expected,
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


def _identity_for_track_item(item: dict[str, Any]) -> OrderIdentity | None:
    """Возвращает identity (kp_id, canonical(plate_name)) для item трека.

    Used by Phase 4: при подсчёте плит по дням нам нужен тот же ключ,
    что и в orders_2d, поэтому plate_name всегда нормализуем.
    """
    kp_id = item.get("kp_id")
    name = item.get("plate_name") or item.get("label") or ""
    canon = _plate_name.canonical(name)
    if kp_id is None or not canon:
        return None
    return (int(kp_id), canon)


def _iter_physical_items(
    track_items: list[dict[str, Any]] | None,
) -> Iterable[dict[str, Any]]:
    """Yields каждый физический item трека: root + все ``secondary_cuts``.

    Каждый ``secondary_cut`` — это отдельная физическая плита,
    полученная из остатка primary-резки. Раньше ``_count_track_items_by_day``
    их не обходил, и identity от secondary терялась → плиты получали
    ``day_number=NULL``. P9: считаем их как полноценные плиты.
    """
    for item in track_items or []:
        if not isinstance(item, dict):
            continue
        yield item
        for sec in item.get("secondary_cuts") or []:
            if isinstance(sec, dict):
                yield sec


def _count_track_items_by_day(
    tracks_by_day: dict[str, list[dict[str, Any]]],
) -> dict[OrderIdentity, dict[int, int]]:
    """Для каждого identity (kp_id, plate_name_canonical) считает qty по day_number.

    Возвращает ``{identity: {day_number: qty}}``. Используется для split'а
    ``mark_plates_as_planned`` по дням, если день у плиты не один.

    P9: учитываются и root items, и ``secondary_cuts`` (они тоже
    самостоятельные физические плиты).
    """
    by_identity: dict[OrderIdentity, dict[int, int]] = defaultdict(
        lambda: defaultdict(int)
    )
    for date_key, day_tracks in tracks_by_day.items():
        for track in day_tracks or []:
            day_number = int(track.get("production_day") or 0)
            if day_number <= 0:
                # production_day могут не проставить — попробуем взять из track-самого
                day_number = int(track.get("day_number") or 0)
            if day_number <= 0:
                continue
            for physical in _iter_physical_items(track.get("items")):
                identity = _identity_for_track_item(physical)
                if identity is None:
                    continue
                by_identity[identity][day_number] += 1
    return {k: dict(v) for k, v in by_identity.items()}


def _covered_weeks_from_tracks(
    tracks_by_day: dict[str, list[dict[str, Any]]] | None,
) -> tuple[date, ...]:
    # Lazy: top-level import of core.production.promise_buckets loads
    # core.production.__init__ → planning → plan_commit (circular).
    from core.production.promise_buckets import iso_week_start

    weeks: set[date] = set()
    for key in tracks_by_day or ():
        try:
            day = date.fromisoformat(str(key)[:10])
        except ValueError:
            continue
        weeks.add(iso_week_start(day))
    return tuple(sorted(weeks))


def _entered_kp_ids(
    orders_with_qty: list[tuple[dict[str, Any], int]],
) -> set[int]:
    entered: set[int] = set()
    for order, qty_to_mark in orders_with_qty:
        kp_id = order.get("kp_id")
        if qty_to_mark > 0 and kp_id:
            entered.add(int(kp_id))
    return entered


def _settle_promises_on_commit(
    *,
    db_path: str,
    plan_id: str,
    entered_kp_ids: set[int],
    covered_weeks: tuple[date, ...],
    settle_fn: PromiseSettleFn | None,
) -> None:
    """Consume / overdue allocations on a connection owned by this commit.

    Plate rows are already marked (existing per-call commits). A settlement
    write failure rolls plates back via ``return_plan_plates_to_production``.
    Overdue is not an error — level 2, commit continues.
    ``settle_fn`` is injected from the app layer (core must not import app).
    """
    if not covered_weeks or settle_fn is None:
        return

    conn = _connect(db_path)
    try:
        conn.execute("PRAGMA foreign_keys = ON")
        settle_fn(
            entered_kp_ids=entered_kp_ids,
            covered_weeks=covered_weeks,
            _external_conn=conn,
        )
        conn.commit()
    except Exception as exc:
        conn.rollback()
        logger.exception(
            "[PLAN_COMMIT] Ошибка погашения обещаний плана %s. Откатываю плиты.",
            plan_id,
        )
        try:
            kp_db_plates.return_plan_plates_to_production(plan_id, db_path)
        except Exception:
            logger.exception(
                "[PLAN_COMMIT] Ошибка при откате плит для плана %s",
                plan_id,
            )
        raise PlanCommitError(
            "Не удалось погасить обещания при коммите плана."
        ) from exc
    finally:
        conn.close()


def commit_plan_plates(
    *,
    plan_id: str,
    orders_2d: list[dict[str, Any]],
    optimization_result: dict[str, Any],
    all_tracks_list: list[dict[str, Any]],
    db_path: str,
    tracks_by_day: dict[str, list[dict[str, Any]]] | None = None,
    day_number_by_date: dict[str, int] | None = None,
    settle_fn: PromiseSettleFn | None = None,
) -> CommitResult:
    """Помечает плиты как «в плане» и валидирует результат.

    Пайплайн:

    1. :func:`count_assigned_plates` → счётчики по источникам.
    2. :func:`distribute_assigned_plates_to_orders` → ``orders_with_qty``.
    3. Валидация: ``optimizer_unmapped`` и ``optimizer_extra`` должны быть
       пустыми (иначе нельзя гарантировать корректность пометки).
    4. Цикл ``kp_db_plates.mark_plates_as_planned`` по каждой позиции заказа.
    5. При любой проблеме — ``kp_db_plates.return_plan_plates_to_production`` и
       выброс :class:`PlanCommitError`.

    Args:
        plan_id: идентификатор плана (используется при записи ``plan_id`` в БД
            и при откате).
        orders_2d: список заказов в виде ``{kp_id, plate_name, qty, ...}``.
        optimization_result: результат оптимизации с ``plate_assignments``.
        all_tracks_list: итоговые дорожки (нужны для учёта RESCUE-плит).
        db_path: путь к ``plita.db``.
        tracks_by_day: ``{date_key: [track, ...]}`` (P5). Если задан,
            ``mark_plates_as_planned`` вызывается per-day, и в каждый
            ``track.items[]`` записывается ``kp_plate_id`` (id строки
            ``kp_plates``). Без этого аргумента — старое поведение.
        day_number_by_date: ``{date_key: day_number}`` (P5). Используется,
            чтобы перевести ``date_key`` в номер дня для записи в БД.
        settle_fn: app-layer callback ``(entered_kp_ids, covered_weeks,
            _external_conn)`` — погашение обещаний в той же tx. Без него
            settle пропускается (core не импортирует app).

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
        # Phase 5 (P8): после Phase 1-4 этот warning не должен срабатывать
        # никогда — identity берётся из единого источника (plate_assignments)
        # и фантомных rescue не возникает. Если всё же сработал — пишем
        # warning и не блокируем commit. Лишние rescue-плиты просто не
        # помечаются в БД (их нет в kp_plates).
        logger.warning(
            "[PLAN_COMMIT] RESCUE-плиты сверх покрытого спроса (safety-net): %s. "
            "Эти плиты не будут помечены в БД, но коммит плана продолжается.",
            rescue_leftovers,
        )

    result = CommitResult(lost_plates=lost_plates)
    marked_any = False

    # P5: если у нас есть tracks_by_day — собираем счётчики qty по дням для
    # каждой identity, чтобы пометить плиты с конкретным day_number и потом
    # записать kp_plate_id в plan.json. Иначе работаем в legacy-режиме.
    counts_by_identity_and_day: dict[OrderIdentity, dict[int, int]] = {}
    if tracks_by_day:
        counts_by_identity_and_day = _count_track_items_by_day(tracks_by_day)

    # P5: пул (plate_id, remaining_qty) пар по identity и дню.
    # ``ids_by_identity_day[(kp_id, canon)][day_number]`` — список пар
    # (plate_id, qty), показывающий, сколько items может разделить
    # каждую запись kp_plates. caller декрементирует remaining_qty при
    # назначении item → kp_plate_id.
    ids_by_identity_day: dict[OrderIdentity, dict[int, list[list[int]]]] = defaultdict(
        lambda: defaultdict(list)
    )

    # Общий mutable бюджет слотов дня на identity: повторные order-строки
    # с тем же (kp_id, canonical name) не должны заново забирать уже
    # израсходованные дни (иначе orphan kp_plates без ссылок с дорожек).
    remaining_by_identity_day: dict[OrderIdentity, dict[int, int]] = {
        ident: dict(days) for ident, days in counts_by_identity_and_day.items()
    }

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

        identity = (int(kp_id), _plate_name.canonical(plate_name))
        # Решение «есть ли дни в треках» — по исходному счётчику; аллокация —
        # только из mutable бюджета (после декремента бюджет может быть 0).
        per_day = counts_by_identity_and_day.get(identity)

        if per_day:
            budget = remaining_by_identity_day.setdefault(identity, {})
            # Распределяем qty_to_mark по оставшемуся бюджету дней.
            # Если сумма бюджета >= qty_to_mark — режем по дням. Если меньше —
            # известные дни из бюджета, остаток без day_number (legacy mismatch).
            total_in_days = sum(budget.values())
            if total_in_days >= qty_to_mark:
                ordered_days = sorted(budget.keys())
                remaining = qty_to_mark
                day_alloc: list[tuple[int, int]] = []
                for d in ordered_days:
                    take = min(budget.get(d, 0), remaining)
                    if take > 0:
                        budget[d] -= take
                        remaining -= take
                        day_alloc.append((d, take))
                    if remaining <= 0:
                        break
                for d, take in day_alloc:
                    mark_result = kp_db_plates.mark_plates_as_planned(
                        kp_id=int(kp_id),
                        plate_name=plate_name,
                        qty_to_plan=take,
                        plan_id=plan_id,
                        db_path=db_path,
                        day_number=d,
                    )
                    _accumulate_mark_result(
                        mark_result,
                        result,
                        kp_id=int(kp_id),
                        plate_name=plate_name,
                        expected=take,
                    )
                    for pid, q in (mark_result.get("id_qty_pairs") or []):
                        ids_by_identity_day[identity][d].append([int(pid), int(q)])
                    if mark_result.get("success") and int(mark_result.get("processed_count", 0) or 0) > 0:
                        marked_any = True
            else:
                # Несоответствие учёта (бюджет дней < qty_to_mark): пишем день
                # для известных слотов, остаток — без дня.
                logger.warning(
                    "[PLAN_COMMIT] Pro-rated plates accounting mismatch: "
                    "identity=%s qty_to_mark=%s sum_per_day=%s",
                    identity, qty_to_mark, total_in_days,
                )
                ordered_days = sorted(budget.keys())
                remaining = qty_to_mark
                day_alloc = []
                for d in ordered_days:
                    take = min(budget.get(d, 0), remaining)
                    if take > 0:
                        budget[d] -= take
                        remaining -= take
                        day_alloc.append((d, take))
                if remaining > 0:
                    day_alloc.append((0, remaining))
                for d, take in day_alloc:
                    if take <= 0:
                        continue
                    mark_result = kp_db_plates.mark_plates_as_planned(
                        kp_id=int(kp_id),
                        plate_name=plate_name,
                        qty_to_plan=take,
                        plan_id=plan_id,
                        db_path=db_path,
                        day_number=d if d > 0 else None,
                    )
                    _accumulate_mark_result(
                        mark_result,
                        result,
                        kp_id=int(kp_id),
                        plate_name=plate_name,
                        expected=take,
                    )
                    if d > 0:
                        for pid, q in (mark_result.get("id_qty_pairs") or []):
                            ids_by_identity_day[identity][d].append([int(pid), int(q)])
                    if mark_result.get("success") and int(mark_result.get("processed_count", 0) or 0) > 0:
                        marked_any = True
        else:
            # Legacy (нет tracks_by_day) — старое поведение без day_number.
            # P9: если tracks_by_day был передан, но per_day для этой
            # identity пуст — это означает, что в track items у плит нет
            # identity (kp_id+plate_name) и backfill_track_items_identity
            # не справился. Пишем явный WARNING — без него такие плиты
            # тихо помечаются с day_number=NULL и зависают вне day_view.
            if tracks_by_day:
                logger.warning(
                    "[PLAN_COMMIT] Identity %s присутствует в "
                    "plate_assignments (qty_to_mark=%s), но в "
                    "tracks_by_day у соответствующих items нет identity. "
                    "Плиты будут помечены БЕЗ day_number и не попадут "
                    "в day_view. Проверить: backfill_track_items_identity "
                    "и наличие kp_id/plate_name у secondary_cuts.",
                    identity, qty_to_mark,
                )
            mark_result = kp_db_plates.mark_plates_as_planned(
                kp_id=int(kp_id),
                plate_name=plate_name,
                qty_to_plan=qty_to_mark,
                plan_id=plan_id,
                db_path=db_path,
            )
            _accumulate_mark_result(
                mark_result,
                result,
                kp_id=int(kp_id),
                plate_name=plate_name,
                expected=qty_to_mark,
            )
            if mark_result.get("success") and int(mark_result.get("processed_count", 0) or 0) > 0:
                marked_any = True

    # P5/P9: записываем kp_plate_id в каждую физическую плиту трека —
    # и в root item, и в каждый secondary_cut. Один plate_id может покрывать
    # несколько физических плит (qty>1) — декрементируем remaining qty в пуле,
    # не удаляя сразу.
    if tracks_by_day:
        for date_key, day_tracks in tracks_by_day.items():
            for track in day_tracks or []:
                day_number = int(track.get("production_day") or 0) or int(track.get("day_number") or 0)
                if day_number <= 0:
                    continue
                for physical in _iter_physical_items(track.get("items")):
                    identity = _identity_for_track_item(physical)
                    if identity is None:
                        continue
                    pool = ids_by_identity_day.get(identity, {}).get(day_number)
                    if not pool:
                        continue
                    pair = next((p for p in pool if p[1] > 0), None)
                    if pair is None:
                        continue
                    physical["kp_plate_id"] = pair[0]
                    pair[1] -= 1

        pool_leftovers: list[dict[str, Any]] = []
        for ident, days in ids_by_identity_day.items():
            for day_num, pool in days.items():
                for pid, rem in pool:
                    if rem > 0:
                        pool_leftovers.append(
                            {
                                "identity": f"{ident[0]}|{ident[1]}",
                                "day": day_num,
                                "plate_id": pid,
                                "remaining": rem,
                            }
                        )
        if pool_leftovers:
            leftover_qty = sum(int(x["remaining"]) for x in pool_leftovers)
            logger.error(
                "[PLAN_COMMIT] После привязки kp_plate_id остались непокрытые "
                "слоты пула (orphan): qty=%s details=%s. Откатываю план %s.",
                leftover_qty,
                pool_leftovers[:30],
                plan_id,
            )
            try:
                kp_db_plates.return_plan_plates_to_production(plan_id, db_path)
            except Exception:
                logger.exception(
                    "[PLAN_COMMIT] Ошибка при откате плит для плана %s",
                    plan_id,
                )
            raise PlanCommitError(
                "После привязки kp_plate_id остались непокрытые слоты пула (orphan)."
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
            kp_db_plates.return_plan_plates_to_production(plan_id, db_path)
        except Exception:
            logger.exception(
                "[PLAN_COMMIT] Ошибка при откате плит для плана %s",
                plan_id,
            )
        raise PlanCommitError(
            f"Не удалось корректно пометить плиты в БД: "
            f"failed={result.plates_failed}, mismatched={result.plates_mismatched}."
        )

    _settle_promises_on_commit(
        db_path=db_path,
        plan_id=plan_id,
        entered_kp_ids=_entered_kp_ids(orders_with_qty),
        covered_weeks=_covered_weeks_from_tracks(tracks_by_day),
        settle_fn=settle_fn,
    )

    return result


__all__ = [
    "CommitResult",
    "PlanCommitError",
    "commit_plan_plates",
    "count_assigned_plates",
    "distribute_assigned_plates_to_orders",
]
