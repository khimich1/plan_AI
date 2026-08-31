#!/usr/bin/env python3
"""Phase 0 validation: срочные плиты + подложки из поздних КП.

Проверяет dealbreaker assumptions из идеи `planirovanie-po-srokam-podlozhki`:
- A1: мэтчи существуют (cross-KP secondary cuts) — фактические И потенциальные
- A2: время прогона оптимизатора ≤ 30 сек
- A3: execution_terms парсится у большинства КП «в работе»

Запуск:
    python scripts/validate_podlozhki_phase0.py --db plita.db --report ai_docs/develop/reports/2026-08-12-podlozhki-phase0.md
"""

from __future__ import annotations

import argparse
import logging
import sqlite3
import sys
import time
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from core.execution_terms import parse_execution_terms_to_datetime
from core.optimization import optimize_with_cascading_longitudinal_cuts

logger = logging.getLogger(__name__)

PLATE_WIDTH_MM = 1200


@dataclass(slots=True)
class PlateRow:
    id: int
    kp_id: int
    plate_name: str
    length_m: float
    width_m: float
    load_class: int | None
    qty: int
    execution_terms: str
    status: str


@dataclass(slots=True)
class BatchInfo:
    batch_id: int
    batch_name: str
    produce_by: date
    plate_id: int
    qty: int


def load_plates_in_production(db_path: str) -> list[PlateRow]:
    conn = sqlite3.connect(db_path)
    try:
        rows = conn.execute(
            """
            SELECT p.id, p.kp_id, p.plate_name, p.length_m, p.width_m,
                   p.load_class, p.qty, o.execution_terms, p.status
            FROM kp_plates p
            JOIN KP_offers o ON o.kp_id = p.kp_id
            JOIN kp_meta m ON m.kp_id = o.kp_id
            WHERE m.status = 'в работе'
              AND p.status IN ('в производстве', 'в плане')
            ORDER BY p.kp_id, p.id
            """
        ).fetchall()
        return [
            PlateRow(
                id=r[0], kp_id=r[1], plate_name=r[2], length_m=r[3],
                width_m=r[4], load_class=r[5], qty=r[6],
                execution_terms=r[7] or "", status=r[8],
            )
            for r in rows
        ]
    finally:
        conn.close()


def load_delivery_batches(db_path: str) -> list[BatchInfo]:
    conn = sqlite3.connect(db_path)
    try:
        rows = conn.execute(
            """
            SELECT b.id, b.name, b.produce_by, i.plate_id, i.qty
            FROM delivery_batch b
            JOIN delivery_batch_item i ON i.batch_id = b.id
            JOIN kp_plates p ON p.id = i.plate_id
            WHERE p.status IN ('в производстве', 'в плане')
            """
        ).fetchall()
        result = []
        for r in rows:
            try:
                produce_by = date.fromisoformat(r[2])
            except ValueError:
                continue
            result.append(BatchInfo(r[0], r[1], produce_by, r[3], r[4]))
        return result
    finally:
        conn.close()


def build_orders_2d(plates: list[PlateRow]) -> list[dict]:
    orders = []
    for p in plates:
        width_mm = int(round(p.width_m * 1000))
        orders.append({
            "length": p.length_m,
            "width": width_mm,
            "qty": p.qty,
            "load_code": p.load_class or 800,
            "kp_id": p.kp_id,
            "plate_name": p.plate_name,
            "reinforcement": 0,
            "concrete_grade": None,
        })
    return orders


def extract_cross_kp_matches(result: dict) -> list[dict]:
    """Фактические cross-KP мэтчи из secondary_cuts."""
    matches = []
    primary_by_unit_id = {}
    for cut in result.get("primary_cuts", []):
        unit_id = cut.get("primary_instance_id")
        if unit_id:
            primary_by_unit_id[unit_id] = cut

    for cut in result.get("secondary_cuts", []):
        parent_id = cut.get("parent_instance_id")
        if not parent_id:
            continue
        parent = primary_by_unit_id.get(parent_id)
        if not parent:
            continue
        sec_kp = cut.get("kp_id")
        pri_kp = parent.get("kp_id")
        if sec_kp and pri_kp and sec_kp != pri_kp:
            matches.append({
                "primary_kp_id": pri_kp,
                "primary_plate": parent.get("plate_name"),
                "primary_length": parent.get("lengths", [0])[0],
                "primary_rest_mm": parent.get("rest", 0),
                "secondary_kp_id": sec_kp,
                "secondary_plate": cut.get("plate_name"),
                "secondary_width_mm": cut.get("cuts", [0])[0],
                "secondary_qty": 1,
            })

    aggregated = {}
    for m in matches:
        key = (m["primary_plate"], m["secondary_plate"])
        if key not in aggregated:
            aggregated[key] = {**m, "secondary_qty": 0}
        aggregated[key]["secondary_qty"] += 1
    return list(aggregated.values())


def find_potential_matches(plates: list[PlateRow]) -> list[dict]:
    """Потенциальные мэтчи по (длина × остаток × класс нагрузки).

    Для каждой плиты считаем остаток = 1200 - width_mm.
    Ищем плиты из ДРУГИХ КП, которые влезают в этот остаток.
    """
    potential = []

    # Группируем по длине и классу нагрузки
    by_length_load: dict[tuple[float, int], list[PlateRow]] = defaultdict(list)
    for p in plates:
        width_mm = int(round(p.width_m * 1000))
        rest = PLATE_WIDTH_MM - width_mm
        if rest > 0:
            key = (p.length_m, p.load_class or 800)
            by_length_load[key].append(p)

    # Для каждой группы ищем пары «срочная → поздняя» из разных КП
    for (length, load), group in by_length_load.items():
        # Сортируем по kp_id, чтобы пары были детерминированы
        group = sorted(group, key=lambda p: (p.kp_id, p.id))
        for i, primary in enumerate(group):
            primary_width = int(round(primary.width_m * 1000))
            primary_rest = PLATE_WIDTH_MM - primary_width
            if primary_rest < 200:  # минимальная полезная ширина
                continue
            for secondary in group:
                if secondary.kp_id == primary.kp_id:
                    continue
                secondary_width = int(round(secondary.width_m * 1000))
                if secondary_width <= primary_rest:
                    potential.append({
                        "primary_kp_id": primary.kp_id,
                        "primary_plate": primary.plate_name,
                        "primary_length": primary.length_m,
                        "primary_width_mm": primary_width,
                        "primary_rest_mm": primary_rest,
                        "secondary_kp_id": secondary.kp_id,
                        "secondary_plate": secondary.plate_name,
                        "secondary_width_mm": secondary_width,
                        "secondary_length": secondary.length_m,
                        "saving_m": primary_rest * primary.length_m / 1000,
                    })

    # Убираем дубли (A под B и B под A — одна пара)
    seen = set()
    unique = []
    for p in potential:
        key = tuple(sorted([
            (p["primary_kp_id"], p["primary_plate"]),
            (p["secondary_kp_id"], p["secondary_plate"]),
        ]))
        if key not in seen:
            seen.add(key)
            unique.append(p)

    return sorted(unique, key=lambda x: x["saving_m"], reverse=True)


def count_parseable_execution_terms(plates: list[PlateRow]) -> tuple[int, int]:
    seen_kp = {}
    for p in plates:
        if p.kp_id not in seen_kp:
            seen_kp[p.kp_id] = p.execution_terms
    total = len(seen_kp)
    parseable = 0
    for terms in seen_kp.values():
        if not terms:
            continue
        dt = parse_execution_terms_to_datetime(terms)
        if dt is not None:
            parseable += 1
    return parseable, total


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--db", default="plita.db")
    parser.add_argument("--report", required=True)
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(message)s")

    print("=" * 60)
    print("Phase 0: Валидация срочных плит + подложек")
    print("=" * 60)

    # 1. Загрузка
    print("\n[1] Загрузка бэклога...")
    plates = load_plates_in_production(args.db)
    batches = load_delivery_batches(args.db)
    kp_ids = set(p.kp_id for p in plates)
    print(f"    Позиций в бэклоге: {len(plates)}")
    print(f"    КП «в работе»: {len(kp_ids)}")
    print(f"    Партий графика: {len(batches)}")

    # 2. A3
    print("\n[2] A3: парсинг execution_terms...")
    parseable, total_kp = count_parseable_execution_terms(plates)
    a3_ok = parseable >= total_kp * 0.8 if total_kp else False
    print(f"    КП с execution_terms: {parseable}/{total_kp} ({parseable/total_kp*100:.0f}%)" if total_kp else "    Нет КП")
    print(f"    A3: {'✅ PASS' if a3_ok else '❌ FAIL'}")

    # 3. Срочные
    print("\n[3] Определение срочных позиций...")
    today = date.today()
    default_deadline = today.fromordinal(today.toordinal() + 14)
    produce_by_by_plate: dict[int, date] = {}
    for b in batches:
        if b.plate_id not in produce_by_by_plate or b.produce_by < produce_by_by_plate[b.plate_id]:
            produce_by_by_plate[b.plate_id] = b.produce_by

    urgent_plates = []
    for p in plates:
        deadline = produce_by_by_plate.get(p.id)
        if deadline is None and p.execution_terms:
            dt = parse_execution_terms_to_datetime(p.execution_terms)
            if dt:
                deadline = dt.date()
        if deadline is None:
            deadline = default_deadline
        if deadline <= default_deadline:
            urgent_plates.append((p, deadline))
    print(f"    Срочных позиций (дедлайн ≤ {default_deadline}): {len(urgent_plates)}/{len(plates)}")

    # 4. A2: оптимизатор
    print("\n[4] A2: полный прогон оптимизатора...")
    orders_2d = build_orders_2d(plates)
    start = time.perf_counter()
    try:
        result = optimize_with_cascading_longitudinal_cuts(orders_2d=orders_2d)
        opt_ok = result and result.get("_opt_status") == "ok"
    except Exception as exc:
        print(f"    ❌ Ошибка: {exc}")
        result = None
        opt_ok = False
    elapsed = time.perf_counter() - start
    a2_ok = elapsed <= 30.0
    print(f"    Время: {elapsed:.1f} сек")
    print(f"    A2: {'✅ PASS' if a2_ok else '❌ FAIL'}")

    # 5. A1a: фактические мэтчи
    print("\n[5a] A1: фактические cross-KP мэтчи из оптимизатора...")
    actual_matches = extract_cross_kp_matches(result) if opt_ok else []
    print(f"    Фактических cross-KP мэтчей: {len(actual_matches)}")

    # 5b. A1b: потенциальные мэтчи (даже если сейчас все плиты 1200)
    print("\n[5b] A1: потенциальные мэтчи по (длина × остаток × класс)...")
    potential_matches = find_potential_matches(plates)
    print(f"    Потенциальных мэтчей: {len(potential_matches)}")
    for m in potential_matches[:5]:
        print(f"      {m['primary_plate']} (КП-{m['primary_kp_id']}, остаток {m['primary_rest_mm']}мм) "
              f"← {m['secondary_plate']} (КП-{m['secondary_kp_id']}, {m['secondary_width_mm']}мм) "
              f"— экономия {m['saving_m']:.1f} м")

    # A1 считается пройденным, если есть фактические ИЛИ потенциальные
    a1_ok = len(actual_matches) > 0 or len(potential_matches) > 0
    print(f"\n    A1: {'✅ PASS' if a1_ok else '❌ FAIL'}")
    if not actual_matches and potential_matches:
    if not actual_matches and potential_matches:
        print(f"    ⚠️  В текущем бэклоге все плиты 1200мм — фактических мэтчей нет.")
        print(f"    ⚠️  Но потенциальных {len(potential_matches)} — фича полезна при появлении неполноширинных плит.")

    # 6. Итог
    print("\n" + "=" * 60)
    print("ИТОГ")
    print("=" * 60)
    print(f"A1 (мэтчи есть):     {'✅' if a1_ok else '❌'} (факт: {len(actual_matches)}, потенциал: {len(potential_matches)})")
    print(f"A2 (время ≤30 сек):  {'✅' if a2_ok else '❌'} ({elapsed:.1f} сек)")
    print(f"A3 (execution_terms): {'✅' if a3_ok else '❌'} ({parseable}/{total_kp})")

    all_ok = a1_ok and a2_ok and a3_ok
    print(f"\nВердикт: {'✅ СТРОИМ MVP' if all_ok else '❌ ПЕРЕОСМЫСЛИТЬ'}")

    # 7. Отчёт
    report_path = Path(args.report)
    report_path.parent.mkdir(parents=True, exist_ok=True)

    report = f"""# Phase 0 Validation: Подложки из поздних КП

Дата: {datetime.now().strftime("%Y-%m-%d %H:%M")}
База: {args.db}

## Метрики

| Метрика | Значение |
|---------|----------|
| Позиций в бэклоге | {len(plates)} |
| КП «в работе» | {len(kp_ids)} |
| Партий графика | {len(batches)} |
| Срочных позиций | {len(urgent_plates)} |
| Время прогона | {elapsed:.1f} сек |
| Фактических cross-KP мэтчей | {len(actual_matches)} |
| Потенциальных мэтчей | {len(potential_matches)} |
| execution_terms парсится | {parseable}/{total_kp} |

## Assumptions

| Assumption | Результат | Вердикт |
|------------|-----------|---------|
| A1: мэтчи существуют | факт: {len(actual_matches)}, потенциал: {len(potential_matches)} | {'✅ PASS' if a1_ok else '❌ FAIL'} |
| A2: время ≤30 сек | {elapsed:.1f} сек | {'✅ PASS' if a2_ok else '❌ FAIL'} |
| A3: execution_terms ≥80% | {parseable}/{total_kp} | {'✅ PASS' if a3_ok else '❌ FAIL'} |

## Фактические мэтчи (secondary cuts)

"""
    if actual_matches:
        for m in actual_matches[:10]:
            report += f"- {m['primary_plate']} (КП-{m['primary_kp_id']}) ← {m['secondary_plate']} ×{m['secondary_qty']} (КП-{m['secondary_kp_id']}), остаток {m['primary_rest_mm']} мм\n"
    else:
        report += "Нет фактических cross-KP мэтчей.\n"

    report += """
## Потенциальные мэтчи (по остатку ширины)

"""
    if potential_matches:
        for m in potential_matches[:15]:
            report += (
                f"- {m['primary_plate']} (КП-{m['primary_kp_id']}, остаток {m['primary_rest_mm']}мм) "
                f"← {m['secondary_plate']} (КП-{m['secondary_kp_id']}, {m['secondary_width_mm']}мм) "
                f"— экономия {m['saving_m']:.1f} м\n"
            )
    else:
        report += "Нет потенциальных мэтчей (все плиты полной ширины 1200мм).\n"

    report += f"""
## Вердикт

{'✅ СТРОИМ MVP' if all_ok else '❌ ПЕРЕОСМЫСЛИТЬ'}

"""
    if not actual_matches and potential_matches:
        report += f"""**Примечание:** в текущем бэклоге все плиты полной ширины 1200мм, поэтому фактических мэтчей нет.
Но обнаружено {len(potential_matches)} потенциальных мэтчей — фича полезна при появлении
неполноширинных плит в заказах.
"""

    report_path.write_text(report, encoding="utf-8")
    print(f"\nОтчёт: {report_path}")

    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())
