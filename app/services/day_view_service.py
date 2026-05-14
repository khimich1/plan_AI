"""Сборка детального вида дня для веб-клиента.

Повторяет логику из `bot/handlers/production_day_view.py::process_day_selection`:
агрегирует дорожки из всех планов на дату и для каждой плиты ищет информацию
в lookup-таблицах с tolerance 0.03 м (fuzzy-поиск).
"""
from __future__ import annotations

import copy
import logging
from typing import Any

from app.core.settings import get_settings
from app.planning import plan_manager
from core import plate_name as plate_name_utils
from core.debug_paths import get_debug_log_path
from core.concrete_grade_resolver import resolve_concrete_grade_from_order

logger = logging.getLogger(__name__)
_DEBUG_AGENT_LOG = get_debug_log_path("debug-ebb546.log")

# P3: tolerance ±0.005 м — защита от float-погрешности.
# Раньше было 0.03 м, что склеивало 5.7 и 5.71 как один заказ и крало identity.
FUZZY_TOLERANCE_M = 0.005


def _reinforcement_to_load_code(reinforcement: float) -> int:
    if reinforcement <= 0:
        return 8
    if reinforcement < 8:
        return 6
    if reinforcement < 12:
        return 8
    if reinforcement < 15:
        return 10
    return 12


def _build_smart_lookup(
    plate_lookup_exact: dict,
    plate_lookup_by_length: dict,
):
    """Создаёт функцию fuzzy-поиска, работающую с копией lookup-таблиц.

    Возвращает запись (словарь) с ключами `customer`, `plate_name`,
    `kp_date`, `kp_id`, `reinforcement` или дефолт, если ничего не найдено.
    Каждый вызов списывает одну плиту (уменьшает `qty_remaining`), поэтому
    lookup должен быть копией.
    """
    formovka_exact = copy.deepcopy(plate_lookup_exact)
    formovka_by_length = copy.deepcopy(plate_lookup_by_length)

    def lookup(length_m: float, width_mm: int) -> dict[str, Any]:
        """Точный поиск identity заказа по (length, width).

        P3: убран опасный fallback width<1200 → (length, 1200), который
        крал identity primary-заказа у вторичного реза.
        Fuzzy сужен до ±0.005 м (защита от float).
        """
        rounded_length = round(length_m, 2)

        key = (rounded_length, width_mm)
        entries = formovka_exact.get(key, [])
        for entry in entries:
            if entry.get("qty_remaining", 0) > 0:
                entry["qty_remaining"] -= 1
                return entry.copy()

        # P3: ТОЛЬКО fuzzy ±0.005 (float-tolerance), без подмены width.
        for lookup_key, entries in formovka_exact.items():
            key_length, key_width = lookup_key
            if key_width != width_mm:
                continue
            if abs(key_length - rounded_length) <= FUZZY_TOLERANCE_M:
                for entry in entries:
                    if entry.get("qty_remaining", 0) > 0:
                        entry["qty_remaining"] -= 1
                        return entry.copy()

        return {
            "kp_id": None,
            "kp_date": "неизвестно",
            "customer": "неизвестно",
            "plate_name": "",
            "reinforcement": 0,
            "concrete_grade": "",
        }

    return lookup


def _iter_plate_items(track: dict):
    """Перебирает основные плиты + вторичные резы (остатки) внутри дорожки.

    Возвращает кортежи `(length_m, width_mm, is_secondary, parent_item, label_hint)`.

    P3: Если у secondary_cut НЕТ ``target_length`` — НЕ берём длину родителя
    (раньше брали, и lookup по (parent_length, sec_width) с fallback width<1200
    крал identity родительского заказа). Без длины secondary_cut пропускается.
    """
    for item in track.get("items", []) or []:
        if item is None:
            continue
        length = item.get("length")
        if not length:
            continue

        mode = item.get("mode", "solid")
        if mode == "transverse" and item.get("width"):
            width_mm = round(item["width"] * 1000)
        elif mode == "split" and item.get("main_w"):
            width_mm = round(item["main_w"] * 1000)
        else:
            width_mm = 1200

        yield float(length), int(width_mm), False, item, None

        for sec in item.get("secondary_cuts") or []:
            sec_width_m = sec.get("width", 0)
            if sec_width_m <= 0:
                continue
            sec_width_mm = round(sec_width_m * 1000)
            sec_length = sec.get("target_length")
            if not sec_length:
                # P3: без явной target_length secondary не должен брать identity
                # родительского заказа — это было главным источником phantom-плит.
                continue
            label_hint = None
            if sec.get("label"):
                label_hint = sec["label"].replace("О ", "").strip()
            yield float(sec_length), int(sec_width_mm), True, item, label_hint


def _aggregate_plates_for_track_from_db(
    track: dict,
    db_rows_by_id: dict[int, dict[str, Any]],
) -> list[dict[str, Any]]:
    """P5: строит plates_info по items с kp_plate_id, читая данные из БД.

    Формат как у :func:`_aggregate_plates_for_track`, плюс учёт строк, исчезнувших
    из ``kp_plates`` после ``complete_day`` (см. :func:`_load_db_rows_for_plan_day`).
    """
    plates: list[dict[str, Any]] = []

    def _add_plate(item: dict[str, Any], is_secondary: bool) -> None:
        plate_id = item.get("kp_plate_id")
        if plate_id is None:
            return
        row = db_rows_by_id.get(int(plate_id))
        if row is None:
            return
        length_m = float(row.get("length_m") or 0)
        width_mm = int(round(float(row.get("width_m") or 0) * 1000))
        load_class = int(row.get("load_class") or 800)
        load_code_value = load_class / 100.0
        load_code: int | float
        if load_code_value.is_integer():
            load_code = int(load_code_value)
        else:
            load_code = load_code_value
        plate_name = row.get("plate_name") or ""
        canon = plate_name_utils.canonical(plate_name)

        existing = next(
            (
                p
                for p in plates
                if p.get("kp_plate_id") == int(plate_id)
            ),
            None,
        )
        if existing:
            existing["qty"] += 1
            return

        plates.append(
            {
                "length_m": round(length_m, 3),
                "width_mm": width_mm,
                "qty": 1,
                "reinforcement": float(row.get("reinforcement") or 0),
                "kp_date": row.get("kp_date") or "неизвестно",
                "customer": row.get("customer") or "неизвестно",
                "kp_id": int(row.get("kp_id") or 0) or None,
                "plate_name": plate_name,
                "plate_name_canonical": canon,
                "load_class": load_class,
                "load_code": load_code,
                "length_dm_raw": row.get("length_dm_raw") or "",
                "is_secondary": bool(is_secondary),
                "kp_plate_id": int(plate_id),
                "write_off_completed": bool(row.get("is_completed_snapshot")),
                "concrete_grade": row.get("concrete_grade") or "",
            }
        )

    for item in track.get("items") or []:
        if not item:
            continue
        _add_plate(item, is_secondary=False)
        for sec in item.get("secondary_cuts") or []:
            if not sec:
                continue
            _add_plate(sec, is_secondary=True)

    plates.sort(key=lambda p: p["length_m"], reverse=True)
    return plates


def _track_has_plate_ids(track: dict) -> bool:
    """True, если у любого item трека есть kp_plate_id (значит план новый)."""
    for item in track.get("items") or []:
        if not item:
            continue
        if item.get("kp_plate_id") is not None:
            return True
        for sec in item.get("secondary_cuts") or []:
            if sec and sec.get("kp_plate_id") is not None:
                return True
    return False


def _load_db_rows_for_plan_day(
    db_path: str, plan_id: str, day_number: int
) -> dict[int, dict[str, Any]]:
    """Загружает данные плит для (plan_id, day_number) из kp_plates и доснимок списанных.

    Живые строки: ``kp_plates`` со status ``в плане``. После ``complete_day`` часть
    id исчезает из ``kp_plates``, но остаётся в ``plate_status_log`` + ``completed_plates`` —
    такие позиции подмешиваем обратно, чтобы веб-список дня не «обнулялся» после списания.
    """
    import sqlite3 as _sql

    def _kp_offer_meta(kp_id: int) -> tuple[str, str]:
        cur.execute(
            """
            SELECT customer_name, execution_terms
            FROM KP_offers
            WHERE kp_id = ?
            LIMIT 1
            """,
            (kp_id,),
        )
        meta = cur.fetchone()
        if meta is None:
            return "неизвестно", "неизвестно"
        customer = meta["customer_name"] if meta["customer_name"] else None
        terms = meta["execution_terms"] if meta["execution_terms"] else None
        return (customer or "неизвестно", terms or "неизвестно")

    rows: dict[int, dict[str, Any]] = {}
    pb_db = str(get_settings().pb_db_path)

    def _cg_for_db_row(
        plate_nm: Any, length_any: Any, load_cls: Any, explicit: Any
    ) -> str:
        s = str(explicit or "").strip()
        if s:
            return s
        return resolve_concrete_grade_from_order(
            {
                "concrete_grade": None,
                "plate_name": plate_nm or "",
                "length": float(length_any or 0) or None,
                "load_code": load_cls or 800,
            },
            db_path=pb_db,
        )

    with _sql.connect(db_path) as conn:
        conn.row_factory = _sql.Row
        cur = conn.cursor()
        cur.execute(
            """
            SELECT p.id, p.kp_id, p.plate_name, p.length_m, p.width_m,
                   p.load_class, p.qty, p.length_dm_raw, p.concrete_grade,
                   k.customer_name, k.execution_terms
            FROM kp_plates p
            LEFT JOIN KP_offers k ON k.kp_id = p.kp_id
            WHERE p.plan_id = ? AND p.day_number = ? AND p.status = 'в плане'
            """,
            (plan_id, day_number),
        )
        for row in cur.fetchall():
            rows[int(row["id"])] = {
                "kp_id": row["kp_id"],
                "plate_name": row["plate_name"],
                "length_m": row["length_m"],
                "width_m": row["width_m"],
                "load_class": row["load_class"],
                "qty": row["qty"],
                "length_dm_raw": row["length_dm_raw"],
                "customer": row["customer_name"],
                "kp_date": row["execution_terms"],
                "reinforcement": 0,
                "is_completed_snapshot": False,
                "concrete_grade": _cg_for_db_row(
                    row["plate_name"],
                    row["length_m"],
                    row["load_class"],
                    row["concrete_grade"],
                ),
            }

        cur.execute(
            """
            SELECT plate_id, kp_id, plate_name, SUM(qty) AS moved_qty
            FROM plate_status_log
            WHERE plan_id = ? AND day_number = ?
              AND to_status = 'completed'
              AND reason = 'completed'
              AND plate_id IS NOT NULL
            GROUP BY plate_id, kp_id, plate_name
            """,
            (plan_id, day_number),
        )
        for slog in cur.fetchall():
            pid = int(slog["plate_id"])
            if pid in rows:
                continue
            ky = int(slog["kp_id"] or 0)
            pname = slog["plate_name"] or ""
            if ky <= 0:
                continue
            cur.execute(
                """
                SELECT length_m, width_m, load_class
                FROM completed_plates
                WHERE kp_id = ? AND plate_name = ? AND production_day = ?
                LIMIT 1
                """,
                (ky, pname, day_number),
            )
            dim = cur.fetchone()
            customer, kp_date = _kp_offer_meta(ky)
            if dim is None:
                rows[pid] = {
                    "kp_id": ky,
                    "plate_name": pname,
                    "length_m": 0.0,
                    "width_m": 0.0,
                    "load_class": 800,
                    "qty": int(slog["moved_qty"] or 0),
                    "length_dm_raw": "",
                    "customer": customer,
                    "kp_date": kp_date,
                    "reinforcement": 0,
                    "is_completed_snapshot": True,
                    "concrete_grade": resolve_concrete_grade_from_order(
                        {
                            "concrete_grade": None,
                            "plate_name": pname or "",
                            "length": None,
                            "load_code": 800,
                        },
                        db_path=pb_db,
                    ),
                }
            else:
                rows[pid] = {
                    "kp_id": ky,
                    "plate_name": pname,
                    "length_m": dim["length_m"],
                    "width_m": dim["width_m"],
                    "load_class": dim["load_class"],
                    "qty": int(slog["moved_qty"] or 0),
                    "length_dm_raw": "",
                    "customer": customer,
                    "kp_date": kp_date,
                    "reinforcement": 0,
                    "is_completed_snapshot": True,
                    "concrete_grade": resolve_concrete_grade_from_order(
                        {
                            "concrete_grade": None,
                            "plate_name": pname or "",
                            "length": dim["length_m"],
                            "load_code": dim["load_class"],
                        },
                        db_path=pb_db,
                    ),
                }
    return rows


def _aggregate_plates_for_track(track: dict, lookup) -> list[dict[str, Any]]:
    plates: list[dict[str, Any]] = []
    is_rescue = track.get("label") == "РЕСКЬЮ"

    for length_m, width_mm, is_secondary, parent_item, label_hint in _iter_plate_items(track):
        info = lookup(length_m, width_mm)

        plate_name = info.get("plate_name") or ""
        if is_rescue and parent_item and (parent_item.get("plate_name") or parent_item.get("label")):
            plate_name = parent_item.get("plate_name") or parent_item.get("label", "")
        if not plate_name and label_hint:
            plate_name = label_hint

        reinforcement = float(info.get("reinforcement") or 0)
        load_code = int(info.get("load_code") or _reinforcement_to_load_code(reinforcement))

        length_dm_raw = info.get("length_dm_raw") or ""
        if not plate_name:
            plate_name = plate_name_utils.make(
                length_m, width_mm, load_code, length_dm_raw=length_dm_raw or None
            )

        # P2: kp_id у secondary без target_length мы уже не подхватываем.
        # У вторичных без identity (label_hint == None) — kp_id остаётся None,
        # такая плита позже сохранится в plate_rests (Фаза 6).
        kp_id = info.get("kp_id")
        if kp_id is None and is_rescue and parent_item is not None:
            kp_id = parent_item.get("kp_id")

        concrete_grade = (
            str(info.get("concrete_grade") or "").strip()
            or (
                str(parent_item.get("concrete_grade") or "").strip()
                if is_rescue and parent_item
                else ""
            )
        )

        # P2: ключ агрегации использует canonical(plate_name) — «Плиты ПБ 45-12-6п»
        # и «ПБ 45-12-6п» считаются одной плитой и больше не дублируются.
        canon = plate_name_utils.canonical(plate_name)
        existing = next(
            (
                p
                for p in plates
                if round(p["length_m"], 2) == round(length_m, 2)
                and p["width_mm"] == width_mm
                and abs(p["reinforcement"] - reinforcement) < 0.1
                and p["kp_date"] == info.get("kp_date", "неизвестно")
                and p["customer"] == info.get("customer", "неизвестно")
                and p.get("kp_id") == kp_id
                and plate_name_utils.canonical(p["plate_name"]) == canon
                and bool(p.get("is_secondary")) == bool(is_secondary)
            ),
            None,
        )
        if existing:
            existing["qty"] += 1
            continue

        plates.append(
            {
                "length_m": round(length_m, 3),
                "width_mm": int(width_mm),
                "qty": 1,
                "reinforcement": reinforcement,
                "kp_date": info.get("kp_date", "неизвестно"),
                "customer": info.get("customer", "неизвестно"),
                "kp_id": kp_id,
                "plate_name": plate_name,
                "load_code": load_code,
                "length_dm_raw": length_dm_raw,
                "is_secondary": bool(is_secondary),
                "concrete_grade": concrete_grade,
            }
        )

    plates.sort(key=lambda p: p["length_m"], reverse=True)
    return plates


def _plan_completion_map(source_plan_ids: list[str], date_key: str) -> dict[str, bool]:
    result: dict[str, bool] = {}
    for plan_id in source_plan_ids:
        plan = plan_manager.load_plan(plan_id)
        if not plan:
            continue
        day = plan.get("days", {}).get(date_key, {})
        result[plan_id] = bool(day.get("completed"))
    return result


def build_day_view_detail(date_key: str, db_path: str | None = None) -> dict | None:
    """Собирает детальный вид дня: дорожки с плитами, сгруппированные по планам.

    P5: для планов, где у items есть ``kp_plate_id`` (новый формат), читает
    plates_info ПРЯМО из БД по ``plan_id+day_number``. Это гарантирует
    инвариант ``plates_info ↔ kp_plates``, и complete_day всегда находит
    нужные строки.

    Для старых (legacy) планов без ``kp_plate_id`` сохраняется fuzzy-lookup
    путь — он помечается флагом ``is_legacy=true`` в каждом track-блоке.

    Возвращает:
        - ``None``, если даты нет ни в одном плане (фронт получит 404);
        - структуру с пустым ``plans`` и ``total_tracks=0``, если дата есть,
          но массив ``tracks`` в плане пустой. Это нужно, чтобы фронт мог
          отличить «дня нет» от «день есть, но без дорожек» и показать
          info-алерт, а не generic-ошибку.
    """
    multi = plan_manager.get_tracks_for_date_from_all_plans(date_key)
    if not multi:
        return None

    tracks: list[dict] = multi.get("tracks") or []
    if not tracks:
        return {
            "date": date_key,
            "plans": [],
            "plans_count": 0,
            "total_tracks": 0,
        }

    lookup = _build_smart_lookup(
        multi.get("plate_lookup_exact", {}),
        multi.get("plate_lookup_by_length", {}),
    )

    source_plans: list[str] = multi.get("source_plans") or []
    completion = _plan_completion_map(source_plans, date_key)

    # Определяем путь по plan'у, чтобы читать БД-данные один раз на (plan_id, day).
    if db_path is None:
        from app.core.settings import get_settings as _get_settings
        db_path = str(_get_settings().plita_db_path)

    db_rows_cache: dict[tuple[str, int], dict[int, dict[str, Any]]] = {}

    plan_blocks: dict[str, dict[str, Any]] = {}
    plan_order: list[str] = []

    for track_index, track in enumerate(tracks, start=1):
        plan_id = track.get("source_plan_id") or "unknown"
        plan_name = track.get("source_plan_name") or plan_id

        block = plan_blocks.get(plan_id)
        if block is None:
            block = {
                "plan_id": plan_id,
                "plan_name": plan_name,
                "completed": completion.get(plan_id, False),
                "tracks": [],
            }
            plan_blocks[plan_id] = block
            plan_order.append(plan_id)

        # P5: пытаемся пойти DB-путём, если у трека есть kp_plate_id.
        is_legacy = True
        plates_info: list[dict[str, Any]] = []
        if _track_has_plate_ids(track):
            day_number = int(track.get("production_day") or 0)
            cache_key = (plan_id, day_number)
            db_rows = db_rows_cache.get(cache_key)
            if db_rows is None and day_number > 0:
                try:
                    db_rows = _load_db_rows_for_plan_day(db_path, plan_id, day_number)
                except Exception:
                    logger.exception(
                        "[DAY_VIEW] DB-путь не сработал для plan=%s day=%s",
                        plan_id, day_number,
                    )
                    db_rows = None
                if db_rows is not None:
                    db_rows_cache[cache_key] = db_rows
            if db_rows is not None:
                plates_info = _aggregate_plates_for_track_from_db(track, db_rows)
                is_legacy = False

        if is_legacy:
            plates_info = _aggregate_plates_for_track(track, lookup)

        block["tracks"].append(
            {
                "track_number": track_index,
                "length": track.get("length"),
                "max_reinforcement": float(track.get("max_reinforcement") or 0),
                "label": track.get("label"),
                "source_plan_id": plan_id,
                "source_plan_name": plan_name,
                "plates_info": plates_info,
                "is_legacy": is_legacy,
            }
        )

    # #region agent log
    try:
        import json as _agent_json
        import time as _agent_time

        _plan_summaries = []
        for _pid, _block in plan_blocks.items():
            _tracks = _block.get("tracks") or []
            _qty_total = 0
            _without_kp = 0
            _legacy_tracks = 0
            _empty_tracks = 0
            for _track in _tracks:
                if _track.get("is_legacy"):
                    _legacy_tracks += 1
                _plates_info = _track.get("plates_info") or []
                if not _plates_info:
                    _empty_tracks += 1
                for _plate in _plates_info:
                    _qty = int(_plate.get("qty") or 0)
                    _qty_total += _qty
                    if not _plate.get("kp_id"):
                        _without_kp += _qty
            _plan_summaries.append({
                "plan_id": _pid,
                "tracks": len(_tracks),
                "legacy_tracks": _legacy_tracks,
                "empty_tracks": _empty_tracks,
                "plates_qty_total": _qty_total,
                "plates_without_kp_qty": _without_kp,
            })
        with open(_DEBUG_AGENT_LOG, "a", encoding="utf-8") as _agent_f:
            _agent_f.write(_agent_json.dumps({
                "sessionId": "ebb546",
                "runId": "pre-fix",
                "hypothesisId": "H3,H4",
                "location": "app/services/day_view_service.py:build_day_view_detail:return",
                "message": "Day-view summary по планам и plates_info",
                "data": {
                    "date_key": date_key,
                    "source_plans": source_plans,
                    "total_tracks": len(tracks),
                    "plans": _plan_summaries,
                },
                "timestamp": int(_agent_time.time() * 1000),
            }, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion

    return {
        "date": date_key,
        "plans": [plan_blocks[pid] for pid in plan_order],
        "plans_count": len(plan_order),
        "total_tracks": len(tracks),
    }
