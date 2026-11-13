#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Разовая сверка очереди «Завести в 1С» (guid_demand) с составом КП.

По умолчанию — сухой прогон: весь путь кода выполняется в транзакции и
откатывается, базы не меняются. Запись — только с ``--apply``.

ЗАПУСКАТЬ ТОЛЬКО НА КОПИИ ПРОДА (``--pb-db`` и ``--plita-db`` с копиями баз).
Первый прогон показывается пользователю до включения сверки на проде.

    .venv/bin/python scripts/reconcile_guid_demand.py \\
        --pb-db /tmp/copy/pb.db --plita-db /tmp/copy/plita.db --report /tmp/report.txt
"""

from __future__ import annotations

import argparse
import sqlite3
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from core.guid_demand_reconcile import format_report, reconcile_guid_demand


def _parse_args(argv: list[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Сверка очереди guid_demand (pb.db) с КП (plita.db). "
            "Без --apply — сухой прогон. Запускать только на копии баз прода."
        )
    )
    parser.add_argument("--pb-db", type=Path, help="путь к pb.db (копия!)")
    parser.add_argument("--plita-db", type=Path, help="путь к plita.db (копия!)")
    parser.add_argument("--apply", action="store_true", help="записать изменения")
    parser.add_argument("--report", type=Path, help="сохранить отчёт в файл")
    return parser.parse_args(argv)


def _resolve_paths(args: argparse.Namespace) -> tuple[Path, Path]:
    pb_path, plita_path = args.pb_db, args.plita_db
    if pb_path is None or plita_path is None:
        from app.core.settings import get_settings

        settings = get_settings()
        pb_path = pb_path or Path(settings.pb_db_path)
        plita_path = plita_path or Path(settings.plita_db_path)
    return Path(pb_path), Path(plita_path)


def _run(pb_path: Path, plita_path: Path, *, apply: bool) -> str:
    pb_conn = sqlite3.connect(str(pb_path), timeout=10)
    plita_conn = sqlite3.connect(f"{plita_path.resolve().as_uri()}?mode=ro", uri=True)
    try:
        pb_conn.execute("BEGIN")
        report = reconcile_guid_demand(pb_conn, plita_conn)
        if apply and report.has_changes:
            pb_conn.commit()
        else:
            pb_conn.rollback()
    finally:
        plita_conn.close()
        pb_conn.close()
    return format_report(report, apply=apply, pb_path=pb_path, plita_path=plita_path)


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)
    pb_path, plita_path = _resolve_paths(args)
    missing = [str(p) for p in (pb_path, plita_path) if not p.is_file()]
    if missing:
        print(f"Базы не найдены: {', '.join(missing)}", file=sys.stderr)
        return 2
    text = _run(pb_path, plita_path, apply=args.apply)
    print(text)
    if args.report:
        args.report.write_text(text + "\n", encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
