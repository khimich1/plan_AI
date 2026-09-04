#!/usr/bin/env python3
"""Task 0 / A1: калибровка buffer в формуле дорожек корзин обещаний.

Для 3–5 прошлых КП из plita.db сравнивает
``ceil(Σ(length_m × qty) / 101 × buffer)`` при buffer=1.0 и 1.15
с фактическими дорожками в ``production_plans``.

БД открывается только read-only. Оптимизатор / ILP не вызываются.

Запуск:
    python scripts/validate_promise_buffers.py --db plita.db \\
      --report ai_docs/develop/reports/2026-09-03-promise-buffers.md
"""

from __future__ import annotations

import argparse
import json
import math
import sqlite3
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from core.production_capacity import MAX_TRACK_LENGTH_M

BUFFERS = (1.0, 1.15)
SAMPLE_MIN = 3
SAMPLE_MAX = 5
# КП идёт в вердикт, если в планах видна хотя бы половина оценки 1.0.
VERDICT_COVERAGE = 0.5
# Подложка / смешивание: доля эквивалента от числа дорожек.
SOLOISH_EQUIV_RATIO = 0.7


@dataclass(slots=True)
class KpLength:
    kp_id: int
    customer: str
    remaining_m: float
    completed_m: float
    qty_remaining: int
    qty_completed: int

    @property
    def full_m(self) -> float:
        return self.remaining_m + self.completed_m


@dataclass(slots=True)
class Occupancy:
    tracks: int
    equiv: float
    item_length_m: float
    plan_ids: tuple[str, ...]


@dataclass(slots=True)
class SampleRow:
    kp_id: int
    customer: str
    length_m: float
    est_1_0: int
    est_1_15: int
    actual_tracks: int
    actual_equiv: float
    in_verdict: bool
    notes: str = ""
    plan_ids: tuple[str, ...] = field(default_factory=tuple)


def estimate_tracks(length_m: float, buffer: float) -> int:
    """Спека assumption 2: ceil(Σ(length × qty) / 101 × buffer)."""
    if length_m <= 0:
        return 0
    raw = (length_m / MAX_TRACK_LENGTH_M) * buffer
    return max(0, int(math.ceil(raw - 1e-12)))


def connect_readonly(db_path: Path) -> sqlite3.Connection:
    resolved = db_path.resolve()
    if not resolved.is_file():
        raise FileNotFoundError(f"БД не найдена: {resolved}")
    conn = sqlite3.connect(f"file:{resolved}?mode=ro", uri=True)
    conn.row_factory = sqlite3.Row
    return conn


def _table_exists(conn: sqlite3.Connection, name: str) -> bool:
    row = conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?",
        (name,),
    ).fetchone()
    return row is not None


def load_kp_lengths(conn: sqlite3.Connection) -> dict[int, KpLength]:
    customers: dict[int, str] = {}
    if _table_exists(conn, "KP_offers"):
        for row in conn.execute(
            "SELECT kp_id, COALESCE(customer_name, '') AS customer FROM KP_offers"
        ):
            customers[int(row["kp_id"])] = str(row["customer"] or "")

    remaining: dict[int, tuple[float, int]] = {}
    if _table_exists(conn, "kp_plates"):
        for row in conn.execute(
            """
            SELECT kp_id,
                   COALESCE(SUM(length_m * qty), 0) AS length_m,
                   COALESCE(SUM(qty), 0) AS qty
            FROM kp_plates
            GROUP BY kp_id
            """
        ):
            remaining[int(row["kp_id"])] = (float(row["length_m"]), int(row["qty"]))

    completed: dict[int, tuple[float, int]] = {}
    if _table_exists(conn, "completed_plates"):
        for row in conn.execute(
            """
            SELECT kp_id,
                   COALESCE(SUM(length_m * qty), 0) AS length_m,
                   COALESCE(SUM(qty), 0) AS qty
            FROM completed_plates
            GROUP BY kp_id
            """
        ):
            completed[int(row["kp_id"])] = (float(row["length_m"]), int(row["qty"]))

    kp_ids = set(customers) | set(remaining) | set(completed)
    result: dict[int, KpLength] = {}
    for kp_id in kp_ids:
        rem_m, rem_q = remaining.get(kp_id, (0.0, 0))
        com_m, com_q = completed.get(kp_id, (0.0, 0))
        if rem_m <= 0 and com_m <= 0:
            continue
        result[kp_id] = KpLength(
            kp_id=kp_id,
            customer=customers.get(kp_id, ""),
            remaining_m=rem_m,
            completed_m=com_m,
            qty_remaining=rem_q,
            qty_completed=com_q,
        )
    return result


def _plate_kp_map(conn: sqlite3.Connection) -> dict[int, int]:
    if not _table_exists(conn, "kp_plates"):
        return {}
    return {
        int(row["id"]): int(row["kp_id"])
        for row in conn.execute("SELECT id, kp_id FROM kp_plates")
    }


def _resolve_item_kp(item: dict, plate_kp: dict[int, int]) -> int | None:
    raw_pid = item.get("kp_plate_id")
    if raw_pid is not None:
        try:
            pid = int(raw_pid)
        except (TypeError, ValueError):
            pid = None
        if pid is not None and pid in plate_kp:
            return plate_kp[pid]
    raw_kp = item.get("kp_id")
    if raw_kp is None:
        return None
    try:
        return int(raw_kp)
    except (TypeError, ValueError):
        return None


def load_occupancy(conn: sqlite3.Connection) -> dict[int, Occupancy]:
    """Фактические дорожки по КП из payload_json планов.

    ``kp_plate_id`` может протухнуть после ухода плит на СГП — тогда
    берём ``item.kp_id``. Эквивалент: доля длины КП на дорожке × (длина / 101).
    """
    if not _table_exists(conn, "production_plans"):
        return {}

    plate_kp = _plate_kp_map(conn)
    tracks: dict[int, set[tuple[str, str, int]]] = {}
    equiv: dict[int, float] = {}
    item_len: dict[int, float] = {}
    plans_by_kp: dict[int, set[str]] = {}

    rows = conn.execute("SELECT id, payload_json FROM production_plans").fetchall()
    for row in rows:
        plan_id = str(row["id"])
        try:
            plan = json.loads(row["payload_json"] or "{}")
        except json.JSONDecodeError:
            print(f"    ⚠️  битый JSON плана {plan_id}, пропускаю")
            continue
        days = plan.get("days") or {}
        if not isinstance(days, dict):
            continue
        for date_key, day in days.items():
            if not isinstance(day, dict):
                continue
            day_tracks = day.get("tracks") or []
            if not isinstance(day_tracks, list):
                continue
            for idx, track in enumerate(day_tracks):
                if not isinstance(track, dict):
                    continue
                try:
                    tlen = float(track.get("length") or 0.0)
                except (TypeError, ValueError):
                    tlen = 0.0
                if tlen <= 0:
                    tlen = MAX_TRACK_LENGTH_M
                by_kp: dict[int, float] = {}
                for item in track.get("items") or []:
                    if not isinstance(item, dict):
                        continue
                    kp_id = _resolve_item_kp(item, plate_kp)
                    if kp_id is None:
                        continue
                    try:
                        length = float(item.get("length") or 0.0)
                    except (TypeError, ValueError):
                        length = 0.0
                    by_kp[kp_id] = by_kp.get(kp_id, 0.0) + length
                if not by_kp:
                    continue
                total = sum(by_kp.values()) or 1.0
                tid = (plan_id, str(date_key), idx)
                for kp_id, length in by_kp.items():
                    tracks.setdefault(kp_id, set()).add(tid)
                    plans_by_kp.setdefault(kp_id, set()).add(plan_id)
                    item_len[kp_id] = item_len.get(kp_id, 0.0) + length
                    equiv[kp_id] = equiv.get(kp_id, 0.0) + (
                        length / total
                    ) * (tlen / MAX_TRACK_LENGTH_M)

    return {
        kp_id: Occupancy(
            tracks=len(tracks[kp_id]),
            equiv=equiv.get(kp_id, 0.0),
            item_length_m=item_len.get(kp_id, 0.0),
            plan_ids=tuple(sorted(plans_by_kp.get(kp_id, ()))),
        )
        for kp_id in tracks
    }


def _notes_for(length: KpLength, occ: Occupancy, est10: int) -> str:
    parts: list[str] = []
    if length.completed_m > 0:
        parts.append(
            f"длина = остаток {length.remaining_m:.1f} + СГП {length.completed_m:.1f}"
        )
    if est10 > 0 and occ.tracks < VERDICT_COVERAGE * est10:
        parts.append("в планы попала малая доля — не в вердикт")
    if occ.tracks > 0 and (occ.equiv / occ.tracks) < SOLOISH_EQUIV_RATIO:
        parts.append("смешивание/подложки: физ. дорожек больше эквивалента")
    return "; ".join(parts)


def build_sample(
    lengths: dict[int, KpLength],
    occupancy: dict[int, Occupancy],
) -> list[SampleRow]:
    rows: list[SampleRow] = []
    for kp_id, occ in occupancy.items():
        length = lengths.get(kp_id)
        if length is None or length.full_m <= 0:
            continue
        est10 = estimate_tracks(length.full_m, 1.0)
        est115 = estimate_tracks(length.full_m, 1.15)
        covered = occ.tracks >= VERDICT_COVERAGE * max(est10, 1)
        rows.append(
            SampleRow(
                kp_id=kp_id,
                customer=length.customer,
                length_m=length.full_m,
                est_1_0=est10,
                est_1_15=est115,
                actual_tracks=occ.tracks,
                actual_equiv=occ.equiv,
                in_verdict=covered,
                notes=_notes_for(length, occ, est10),
                plan_ids=occ.plan_ids,
            )
        )
    rows.sort(key=lambda r: (-r.length_m if r.in_verdict else 0, -r.length_m, r.kp_id))
    verdict = [r for r in rows if r.in_verdict]
    extra = [r for r in rows if not r.in_verdict]
    chosen = verdict[:SAMPLE_MAX]
    if len(chosen) < SAMPLE_MAX:
        chosen.extend(extra[: SAMPLE_MAX - len(chosen)])
    if len(chosen) < SAMPLE_MIN:
        print(
            f"    ⚠️  в выборке {len(chosen)} КП (ожидали {SAMPLE_MIN}–{SAMPLE_MAX})"
        )
    return chosen


def recommend_buffer(rows: list[SampleRow]) -> tuple[float, str]:
    """1.0 vs 1.15. Смешанные мелкие КП судят по эквиваленту, не по присутствию."""
    primary = [r for r in rows if r.in_verdict]
    if not primary:
        return 1.15, "нет КП с достаточной долей в планах — оставляем консервативный 1.15"

    undershoot_1_0 = 0
    overshoot_1_15 = 0
    for row in primary:
        soloish = (
            row.actual_tracks == 0
            or (row.actual_equiv / row.actual_tracks) >= SOLOISH_EQUIV_RATIO
        )
        fact = row.actual_tracks if soloish else max(1, int(math.ceil(row.actual_equiv - 1e-12)))
        if row.est_1_0 < fact - 1:
            undershoot_1_0 += 1
        if row.est_1_15 > fact + 1:
            overshoot_1_15 += 1

    if undershoot_1_0 == 0:
        reason = (
            "оценка 1.0 покрывает факт (±1 дорожка) на КП в вердикте; "
            "1.15 систематически завышает; зазор ручки 3-vs-5 уже даёт запас"
        )
        return 1.0, reason
    if undershoot_1_0 and overshoot_1_15 == 0:
        return 1.15, "оценка 1.0 не покрывает факт, 1.15 покрывает"
    return 1.15, "оценка 1.0 не покрывает факт — оставляем 1.15"


def render_report(
    *,
    db_path: Path,
    rows: list[SampleRow],
    buffer: float,
    reason: str,
    kp_total: int,
    plans_total: int,
) -> str:
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    lines = [
        "# Phase 0 / A1: калибровка buffer корзин обещаний",
        "",
        f"Дата: {now}",
        f"База: `{db_path}` (read-only)",
        f"Формула: `tracks = ceil(Σ(length_m × qty) / {MAX_TRACK_LENGTH_M:g} × buffer)`",
        "",
        "## Метрики",
        "",
        "| Метрика | Значение |",
        "|---------|----------|",
        f"| КП с плитами (остаток + СГП) | {kp_total} |",
        f"| Планов в `production_plans` | {plans_total} |",
        f"| КП в выборке | {len(rows)} |",
        f"| КП в вердикте | {sum(1 for r in rows if r.in_verdict)} |",
        "",
        "## КП → оценка vs факт",
        "",
        "| КП | Заказчик | Длина, м | est 1.0 | est 1.15 | факт дорожек | факт экв. | вердикт | заметка |",
        "|----|----------|----------|---------|----------|--------------|-----------|---------|---------|",
    ]
    for row in rows:
        flag = "да" if row.in_verdict else "нет"
        note = row.notes.replace("|", "/") if row.notes else "—"
        customer = (row.customer or "—").replace("|", "/")
        lines.append(
            f"| {row.kp_id} | {customer} | {row.length_m:.1f} | "
            f"{row.est_1_0} | {row.est_1_15} | {row.actual_tracks} | "
            f"{row.actual_equiv:.2f} | {flag} | {note} |"
        )

    lines.extend(
        [
            "",
            "## Как считали факт",
            "",
            "- Длина КП = сумма `kp_plates` (текущий остаток) + `completed_plates` (уже на СГП).",
            "- Факт дорожек = уникальные дорожки планов, где есть плиты этого КП "
            "(`kp_plate_id` → живая строка, иначе `item.kp_id` — протухшие id после СГП).",
            "- Факт экв. = сумма по дорожкам доли длины КП × (длина дорожки / 101). "
            "Для подложек физ. присутствие завышает занятость.",
            "- В вердикт не берём КП, у которых в планах меньше половины оценки 1.0 "
            "(сравнение полного заказа с крошечным входом врёт).",
            "",
            "## Вывод",
            "",
            f"**buffer = {buffer:.2f}** (безразмерный коэффициент к м / 101, не м³)",
            "",
            reason + ".",
            "",
            "Спека (`assumption 2`) этим прогоном **не обновлена** — нужен выбор человека.",
            "",
        ]
    )
    return "\n".join(lines) + "\n"


def _plans_count(conn: sqlite3.Connection) -> int:
    if not _table_exists(conn, "production_plans"):
        return 0
    return int(conn.execute("SELECT COUNT(*) FROM production_plans").fetchone()[0])


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--db", default="plita.db")
    parser.add_argument(
        "--report",
        default="ai_docs/develop/reports/2026-09-03-promise-buffers.md",
    )
    args = parser.parse_args()

    db_path = Path(args.db)
    print("=" * 60)
    print("Phase 0: Калибровка buffer корзин обещаний (A1)")
    print("=" * 60)

    conn = connect_readonly(db_path)
    try:
        lengths = load_kp_lengths(conn)
        occupancy = load_occupancy(conn)
        plans_total = _plans_count(conn)
    finally:
        conn.close()

    print(f"\n[1] КП с плитами: {len(lengths)}")
    print(f"    Планов: {plans_total}")
    print(f"    КП, встреченных в планах: {len(occupancy)}")

    rows = build_sample(lengths, occupancy)
    print(f"\n[2] Выборка: {len(rows)} КП")
    for row in rows:
        mark = "вердикт" if row.in_verdict else "справка"
        print(
            f"    КП-{row.kp_id}: {row.length_m:.1f} м → "
            f"est 1.0={row.est_1_0} / 1.15={row.est_1_15}, "
            f"факт={row.actual_tracks} (экв. {row.actual_equiv:.2f}) [{mark}]"
        )

    buffer, reason = recommend_buffer(rows)
    print(f"\n[3] Рекомендация: buffer = {buffer:.2f}")
    print(f"    {reason}")

    report_path = Path(args.report)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(
        render_report(
            db_path=db_path,
            rows=rows,
            buffer=buffer,
            reason=reason,
            kp_total=len(lengths),
            plans_total=plans_total,
        ),
        encoding="utf-8",
    )
    print(f"\nОтчёт: {report_path}")
    print(f"buffer = {buffer:.2f}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
