#!/usr/bin/env python3
"""Регрессионная проверка потери плит при планировании на фикстуре заказа.

Запуск из корня репозитория:
    python scripts/run_plate_loss_regression.py

Отчёт: reports/plate_loss_regression_<timestamp>.md
"""
from __future__ import annotations

import json
import re
import sqlite3
import sys
import time
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from core import kp_db
from core.config_and_data import canonical_plate_key
from core.optimization import optimize_with_cascading_longitudinal_cuts, verify_coverage
from core.optimization.result_contract import is_optimization_success
from core.plate_line_parser import parse_line
from core.plate_name import canonical, display
from core.plan_commit import (
    commit_plan_plates,
    count_assigned_plates,
    distribute_assigned_plates_to_orders,
)
from core.production.dto import (
    LoadConfig,
    OptimizeConfig,
    PersistConfig,
    PlanBuildInput,
)
from core.production.planning import load, optimize, persist

DEFAULT_FIXTURE = PROJECT_ROOT / "tests/fixtures/regression/roman_20260503_plates.txt"
KP_ID = 999
PLAN_START = "2026-05-05"
TRACKS_PER_DAY = 3


@dataclass
class ParsedLine:
    line_no: int
    raw: str
    plate_name: str
    length_m: float
    width_m: float
    load_code: float
    qty: int
    length_dm_raw: str = ""


@dataclass
class RegressionReport:
    fixture_path: str
    started_at: str
    duration_sec: float = 0.0
    total_lines: int = 0
    parsed_lines: int = 0
    unparsed_lines: list[str] = field(default_factory=list)
    unparsed_qty_estimate: int = 0
    total_qty_input: int = 0
    total_qty_parsed: int = 0
    unique_positions: int = 0
    optimizer_ok: bool | None = None
    verify_coverage_ok: bool | None = None
    verify_missing: dict = field(default_factory=dict)
    lost_plates_commit: list[dict] = field(default_factory=list)
    plates_skipped: int = 0
    plates_marked: int = 0
    unmapped_assignments: int = 0
    rescue_tracks: int = 0
    total_tracks: int = 0
    total_days: int = 0
    db_before: int = 0
    db_in_plan: int = 0
    db_still_production: int = 0
    db_balance_ok: bool | None = None
    full_coverage_ok: bool | None = None
    plan_build_error: str | None = None
    per_name_gaps: list[dict] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)


def _extract_plate_name(raw_line: str) -> str:
    s = raw_line.strip()
    s = re.sub(r"\s+\d+\s*$", "", s).strip()
    if not s.lower().startswith("плиты "):
        s = display(s)
    return s


def _parse_fixture(path: Path) -> tuple[list[ParsedLine], list[str]]:
    parsed: list[ParsedLine] = []
    unparsed: list[str] = []
    lines = [
        ln.strip()
        for ln in path.read_text(encoding="utf-8").splitlines()
        if ln.strip() and not ln.strip().startswith("#")
    ]
    for idx, raw in enumerate(lines, start=1):
        result = parse_line(raw)
        if not result.parsed:
            unparsed.append(f"L{idx}: {raw} ({result.reason_text or result.reason_code})")
            continue
        load_code = float(result.load_code or 8)
        parsed.append(
            ParsedLine(
                line_no=idx,
                raw=raw,
                plate_name=_extract_plate_name(raw),
                length_m=result.length_m,
                width_m=result.width_m,
                load_code=load_code,
                qty=int(result.qty or 1),
                length_dm_raw=result.length_dm_raw or "",
            )
        )
    return parsed, unparsed


def _seed_db(db_path: str, items: list[ParsedLine]) -> int:
    kp_db.init_schema(db_path)
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "INSERT INTO KP_offers (kp_id, creation_date, execution_terms, customer_name) "
            "VALUES (?, '2026-05-03', '01.06.2026', 'Regression Roman 20260503')",
            (KP_ID,),
        )
        conn.execute("INSERT INTO kp_meta (kp_id, status) VALUES (?, 'в работе')", (KP_ID,))
        pos = 1
        for item in items:
            width_mm = int(round(item.width_m * 1000))
            load_class = int(round(item.load_code * 100))
            conn.execute(
                """
                INSERT INTO kp_plates (
                    kp_id, position_number, plate_name, length_m, width_m,
                    load_class, qty, status
                ) VALUES (?, ?, ?, ?, ?, ?, ?, 'в производстве')
                """,
                (
                    KP_ID,
                    pos,
                    item.plate_name,
                    item.length_m,
                    item.width_m,
                    load_class,
                    item.qty,
                ),
            )
            pos += 1
        conn.commit()
    return pos - 1


def _db_totals(db_path: str, *, plan_id: str | None = None) -> tuple[int, int, dict, dict]:
    with sqlite3.connect(db_path) as conn:
        in_prod_rows = conn.execute(
            "SELECT plate_name, SUM(qty) FROM kp_plates "
            "WHERE kp_id=? AND status='в производстве' GROUP BY plate_name",
            (KP_ID,),
        ).fetchall()
        in_plan_rows = conn.execute(
            "SELECT plate_name, SUM(qty) FROM kp_plates "
            "WHERE kp_id=? AND status='в плане' AND plan_id=? GROUP BY plate_name",
            (KP_ID, plan_id),
        ).fetchall() if plan_id else []
    in_prod = {name: int(qty) for name, qty in in_prod_rows}
    in_plan = {name: int(qty) for name, qty in in_plan_rows}
    return sum(in_prod.values()), sum(in_plan.values()), in_prod, in_plan


def _demand_from_orders(orders_2d: list[dict]) -> dict:
    demand: dict = {}
    for order in orders_2d:
        key = canonical_plate_key(order["length"], order["width"], order.get("load_code", 8))
        demand[key] = demand.get(key, 0) + int(order.get("qty", 1))
    return demand


def _count_plan_tracks(plan: dict) -> int:
    total = 0
    for day in (plan.get("days") or {}).values():
        total += len(day.get("tracks") or [])
    return total


def run_regression(fixture_path: Path, pb_db_path: Path) -> RegressionReport:
    report = RegressionReport(
        fixture_path=str(fixture_path),
        started_at=datetime.now().isoformat(timespec="seconds"),
    )
    t0 = time.monotonic()

    if not fixture_path.exists():
        report.errors.append(f"Fixture not found: {fixture_path}")
        report.duration_sec = time.monotonic() - t0
        return report

    parsed, unparsed = _parse_fixture(fixture_path)
    report.total_lines = len(parsed) + len(unparsed)
    report.parsed_lines = len(parsed)
    report.unparsed_lines = unparsed
    report.total_qty_parsed = sum(p.qty for p in parsed)
    report.total_qty_input = report.total_qty_parsed  # unparsed not counted
    report.unique_positions = len(parsed)

    if not parsed:
        report.errors.append("No parsed plates — cannot continue.")
        report.duration_sec = time.monotonic() - t0
        return report

    tmp_dir = PROJECT_ROOT / "tmp" / "plate_loss_regression"
    tmp_dir.mkdir(parents=True, exist_ok=True)
    db_path = str(tmp_dir / "plita_regression.db")
    Path(db_path).unlink(missing_ok=True)

    _seed_db(db_path, parsed)
    report.db_before, _, before_by_name, _ = _db_totals(db_path)

    from app.repositories.plan_repository import PlanRepository
    from app.services.plan_distribution_service import PlanLoadAdapter, PlanPersistAdapter
    from app.services.plan_distribution_service import PlanDistributionService

    plan_input = PlanBuildInput(
        start_date=PLAN_START,
        tracks_count=TRACKS_PER_DAY,
        filter_method="kp",
        selected_kp_ids=(KP_ID,),
    )
    load_config = LoadConfig(plita_db_path=db_path, pb_db_path=str(pb_db_path))
    plan_load = PlanLoadAdapter(PlanRepository(db_path=db_path))

    try:
        load_result = load(plan_input, config=load_config, plan_load=plan_load)
    except Exception as exc:
        report.plan_build_error = f"load failed: {exc}"
        report.errors.append(report.plan_build_error)
        report.duration_sec = time.monotonic() - t0
        return report

    orders_2d = load_result.orders_2d
    orders_qty = sum(int(o.get("qty", 0)) for o in orders_2d)
    if orders_qty != report.db_before:
        report.errors.append(
            f"load qty mismatch: orders_2d={orders_qty}, db={report.db_before}"
        )

    # --- Phase A: optimizer ---
    try:
        opt_result = optimize(
            load_result,
            config=OptimizeConfig(pb_db_path=str(pb_db_path), layout_reinforcement_order="asc"),
        )
    except Exception as exc:
        report.plan_build_error = f"optimize failed: {exc}"
        report.errors.append(report.plan_build_error)
        report.duration_sec = time.monotonic() - t0
        return report

    optimization_result = opt_result.optimization_result or {}
    report.optimizer_ok = is_optimization_success(optimization_result)
    report.total_tracks = len(opt_result.all_tracks_list)
    report.rescue_tracks = sum(
        1 for t in opt_result.all_tracks_list if str(t.get("label") or "").upper() == "РЕСКЬЮ"
    )

    demand = _demand_from_orders(orders_2d)
    cov = verify_coverage(
        demand,
        optimization_result.get("primary_cuts", []) or [],
        optimization_result.get("secondary_cuts", []) or [],
    )
    report.verify_coverage_ok = bool(cov.get("ok"))
    report.verify_missing = dict(cov.get("missing") or {})

    assigned, unmapped = count_assigned_plates(
        optimization_result, opt_result.all_tracks_list
    )
    report.unmapped_assignments = sum(len(v) for v in unmapped.values())
    lost, orders_with_qty, _leftovers = distribute_assigned_plates_to_orders(
        orders_2d, assigned
    )
    report.lost_plates_commit = lost
    report.plates_skipped = sum(
        1
        for order, qty_to_mark in orders_with_qty
        if qty_to_mark <= 0 and int(order.get("qty", 0) or 0) > 0
    )

    # --- Phase B: full persist ---
    try:
        repo = PlanRepository(db_path=db_path)
        distribution = PlanDistributionService()
        persist_port = PlanPersistAdapter(repo, distribution)
        persist_result = persist(
            load_result,
            opt_result,
            PersistConfig(
                plita_db_path=db_path,
                start_date=PLAN_START,
                tracks_count=TRACKS_PER_DAY,
                layout_reinforcement_order="asc",
            ),
            persist_port,
        )
        plan = persist_result.plan
        report.total_days = len(plan.get("days") or {})
        report.total_tracks = _count_plan_tracks(plan)
        plan_id = plan["id"]
        report.plates_marked = report.db_before - (
            _db_totals(db_path, plan_id=plan_id)[0]
        )

    except Exception as exc:
        report.plan_build_error = f"persist failed: {exc}"
        report.errors.append(report.plan_build_error)
        report.duration_sec = time.monotonic() - t0
        return report

    still_prod, in_plan, still_by_name, in_plan_by_name = _db_totals(db_path, plan_id=plan_id)
    report.db_in_plan = in_plan
    report.db_still_production = still_prod
    report.db_balance_ok = report.db_before == in_plan + still_prod
    report.full_coverage_ok = still_prod == 0 and not report.lost_plates_commit

    # Per plate_name gaps (canonical)
    input_by_name: Counter[str] = Counter()
    for item in parsed:
        input_by_name[canonical(item.plate_name)] += item.qty

    after_by_name: Counter[str] = Counter()
    for name, qty in in_plan_by_name.items():
        after_by_name[canonical(name)] += qty
    for name, qty in still_by_name.items():
        after_by_name[canonical(name)] += qty

    for name, qty_in in sorted(input_by_name.items()):
        accounted = after_by_name.get(name, 0)
        if accounted < qty_in:
            report.per_name_gaps.append(
                {
                    "plate_name": name,
                    "input_qty": qty_in,
                    "accounted_qty": accounted,
                    "missing_qty": qty_in - accounted,
                }
            )
        marked = sum(
            int(qty)
            for n, qty in in_plan_by_name.items()
            if canonical(n) == name
        )
        left = sum(
            int(qty)
            for n, qty in still_by_name.items()
            if canonical(n) == name
        )
        if left > 0:
            report.per_name_gaps.append(
                {
                    "plate_name": name,
                    "input_qty": qty_in,
                    "in_plan_qty": marked,
                    "still_production_qty": left,
                    "not_in_plan_qty": left,
                }
            )

    report.duration_sec = time.monotonic() - t0
    return report


def _md_report(r: RegressionReport) -> str:
    lines = [
        "# Отчёт: регрессия потери плит",
        "",
        f"- **Фикстура:** `{r.fixture_path}`",
        f"- **Запуск:** {r.started_at}",
        f"- **Длительность:** {r.duration_sec:.1f} с",
        "",
        "## 1. Парсинг входного текста",
        "",
        f"| Метрика | Значение |",
        f"|---------|----------|",
        f"| Строк (всего) | {r.total_lines} |",
        f"| Распознано строк | {r.parsed_lines} |",
        f"| Не распознано | {len(r.unparsed_lines)} |",
        f"| Плит (qty) в распознанных | {r.total_qty_parsed} |",
        f"| Уникальных позиций в БД | {r.unique_positions} |",
        "",
    ]

    if r.unparsed_lines:
        lines += ["### Нераспознанные строки", ""]
        for item in r.unparsed_lines[:30]:
            lines.append(f"- {item}")
        if len(r.unparsed_lines) > 30:
            lines.append(f"- … и ещё {len(r.unparsed_lines) - 30}")
        lines.append("")

    lines += [
        "## 2. Оптимизатор (уровень A)",
        "",
        f"- **optimizer success:** {r.optimizer_ok}",
        f"- **verify_coverage.ok:** {r.verify_coverage_ok}",
        f"- **verify_coverage missing keys:** {len(r.verify_missing)}",
        f"- **unmapped assignments:** {r.unmapped_assignments}",
        f"- **дорожек (вкл. RESCUE):** {r.total_tracks} (RESCUE: {r.rescue_tracks})",
        "",
    ]
    if r.verify_missing:
        lines += ["### verify_coverage missing (sample)", "", "```json"]
        lines.append(json.dumps(dict(list(r.verify_missing.items())[:20]), ensure_ascii=False, indent=2))
        lines.append("```")
        lines.append("")

    lines += [
        "## 3. Commit / план (уровень B)",
        "",
        f"- **lost_plates (commit):** {len(r.lost_plates_commit)} позиций",
        f"- **plates_skipped:** {r.plates_skipped}",
        f"- **plates_marked (повторный commit):** {r.plates_marked}",
        f"- **дней в плане:** {r.total_days}",
        "",
    ]
    if r.lost_plates_commit:
        lines += ["### lost_plates", "", "```json"]
        lines.append(json.dumps(r.lost_plates_commit[:30], ensure_ascii=False, indent=2))
        lines.append("```")
        lines.append("")

    lines += [
        "## 4. Баланс БД (уровень C)",
        "",
        f"| | qty |",
        f"|--|-----|",
        f"| До планирования («в производстве») | {r.db_before} |",
        f"| После: «в плане» | {r.db_in_plan} |",
        f"| После: осталось «в производстве» | {r.db_still_production} |",
        f"| **Баланс (не исчезли из учёта)** | **{'OK' if r.db_balance_ok else 'FAIL'}** |",
        f"| **100% в плане** | **{'OK' if r.full_coverage_ok else 'НЕТ'}** |",
        "",
    ]

    if r.per_name_gaps:
        lines += ["### Позиции с расхождениями", ""]
        for gap in r.per_name_gaps[:40]:
            lines.append(f"- `{gap['plate_name']}`: {gap}")
        if len(r.per_name_gaps) > 40:
            lines.append(f"- … и ещё {len(r.per_name_gaps) - 40}")
        lines.append("")

    if r.plan_build_error or r.errors:
        lines += ["## Ошибки", ""]
        for err in r.errors:
            lines.append(f"- {err}")
        if r.plan_build_error:
            lines.append(f"- {r.plan_build_error}")
        lines.append("")

    verdict = "PASS"
    if r.errors or r.plan_build_error:
        verdict = "FAIL (ошибка пайплайна)"
    elif not r.db_balance_ok:
        verdict = "FAIL (плиты исчезли из учёта)"
    elif r.lost_plates_commit or r.db_still_production > 0 or not r.verify_coverage_ok:
        verdict = "WARN (неполное покрытие, плиты в производстве)"
    elif r.unparsed_lines:
        verdict = "WARN (часть строк не распознана парсером)"

    lines += [
        "## Вердикт",
        "",
        f"**{verdict}**",
        "",
    ]
    return "\n".join(lines)


def main() -> int:
    fixture = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_FIXTURE
    pb_db = PROJECT_ROOT / "docker" / "seed" / "pb.db"
    if not pb_db.exists():
        print(f"pb.db not found at {pb_db}", file=sys.stderr)
        return 2

    print(f"Running plate loss regression on {fixture} ...")
    report = run_regression(fixture, pb_db)
    md = _md_report(report)

    reports_dir = PROJECT_ROOT / "reports"
    reports_dir.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = reports_dir / f"plate_loss_regression_{stamp}.md"
    out_path.write_text(md, encoding="utf-8")

    print(md)
    print(f"\nReport saved: {out_path}")
    return 0 if report.db_balance_ok and not report.errors else 1


if __name__ == "__main__":
    raise SystemExit(main())
