#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Метрика пилота «Логистика»: доля рейсов, закрытых без ручной правки состава (SHIP-603).

Сравнивает ``shipments.propose_snapshot`` (JSON, что предложила система) с финальным
составом ``shipment_items`` у done-рейсов. Совпадение — по мультимножеству строк
``(completed_plate_id | mark, qty)``: правки веса/заметок/порядка укладки правкой
состава не считаются.

Бакеты:
  - match        — состав совпал со снимком propose (правок не было);
  - edited       — состав изменён логистом (qty/строки добавлены/убраны);
  - no_snapshot  — propose ни разу не вызывался (снимка нет).

Gate пилота из спеки: hit-rate ≥ 50% среди done-рейсов со снимком.

Запуск из корня репозитория:
    ./.venv/bin/python scripts/shipment_propose_hitrate.py [--db plita.db] [--verbose]
"""

from __future__ import annotations

import argparse
import json
import sqlite3
import sys
from collections import Counter
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from core.kp_db_common import DEFAULT_DB

HIT_RATE_GATE_PERCENT = 50.0


def _item_key(item_type: str, completed_plate_id: int | None, mark: str | None) -> tuple:
    if item_type == "plate":
        return ("plate", completed_plate_id)
    return ("free", (mark or "").strip())


def _snapshot_multiset(snapshot: dict) -> Counter:
    rows = Counter()
    for item in snapshot.get("items") or []:
        key = _item_key(
            str(item.get("item_type") or "plate"),
            item.get("completed_plate_id"),
            item.get("mark"),
        )
        rows[key] += int(item.get("qty") or 0)
    return rows


def _final_multiset(cur: sqlite3.Cursor, shipment_id: int) -> Counter:
    cur.execute(
        "SELECT item_type, completed_plate_id, mark, qty FROM shipment_items WHERE shipment_id = ?",
        (shipment_id,),
    )
    rows = Counter()
    for item_type, completed_plate_id, mark, qty in cur.fetchall():
        rows[_item_key(str(item_type), completed_plate_id, mark)] += int(qty or 0)
    return rows


def collect_stats(db_path: str) -> dict:
    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT id, shipment_date, propose_snapshot FROM shipments WHERE status = 'done' ORDER BY id"
        )
        done = cur.fetchall()
        stats = {"done_total": len(done), "match": [], "edited": [], "no_snapshot": [], "invalid_snapshot": []}
        for shipment_id, shipment_date, snapshot_raw in done:
            entry = {"shipment_id": shipment_id, "shipment_date": shipment_date}
            if not snapshot_raw:
                stats["no_snapshot"].append(entry)
                continue
            try:
                snapshot = json.loads(snapshot_raw)
            except (TypeError, ValueError):
                stats["invalid_snapshot"].append(entry)
                continue
            proposed = _snapshot_multiset(snapshot)
            final = _final_multiset(cur, int(shipment_id))
            bucket = "match" if proposed == final else "edited"
            stats[bucket].append(
                {**entry, "proposed": sum(proposed.values()), "final": sum(final.values())}
            )
        return stats


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Доля done-рейсов, закрытых без ручной правки состава propose."
    )
    parser.add_argument(
        "--db",
        default=DEFAULT_DB,
        help=f"Путь к plita.db (по умолчанию: {DEFAULT_DB})",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Печатать разбор по каждому рейсу",
    )
    args = parser.parse_args()

    db_path = Path(args.db)
    if not db_path.is_file():
        print(f"❌ База не найдена: {db_path}")
        return 1

    try:
        stats = collect_stats(str(db_path))
    except sqlite3.OperationalError as exc:
        print(f"❌ Не удалось прочитать shipments: {exc} (схема логистики ещё не создана?)")
        return 1

    done_total = stats["done_total"]
    with_snapshot = len(stats["match"]) + len(stats["edited"]) + len(stats["invalid_snapshot"])
    comparable = len(stats["match"]) + len(stats["edited"])
    hit_rate = (len(stats["match"]) / comparable * 100.0) if comparable else None

    print("== Propose hit-rate (done-рейсы) ==")
    print(f"Done-рейсов всего:            {done_total}")
    print(f"  без снимка propose:         {len(stats['no_snapshot'])}")
    if stats["invalid_snapshot"]:
        print(f"  битый снимок (вне метрики): {len(stats['invalid_snapshot'])}")
    print(f"  со снимком:                 {with_snapshot}")
    print(f"    match (без правки):       {len(stats['match'])}")
    print(f"    edited (ручная правка):   {len(stats['edited'])}")

    if args.verbose:
        for bucket in ("match", "edited", "no_snapshot", "invalid_snapshot"):
            if not stats[bucket]:
                continue
            print(f"\n-- {bucket} --")
            for entry in stats[bucket]:
                line = f"  рейс #{entry['shipment_id']} от {entry['shipment_date']}"
                if "proposed" in entry:
                    line += f" (предложено {entry['proposed']} шт → отгружено {entry['final']} шт)"
                print(line)

    if hit_rate is None:
        print("\nHit-rate: н/д (нет done-рейсов со снимком propose)")
        return 0

    verdict = "PASS" if hit_rate >= HIT_RATE_GATE_PERCENT else "BELOW GATE"
    print(f"\nHit-rate: {hit_rate:.1f}% (gate пилота ≥ {HIT_RATE_GATE_PERCENT:.0f}%) — {verdict}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
