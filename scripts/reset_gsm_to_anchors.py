#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Сброс ГСМ к импортированным якорям (подготовка к тестовому прогону).

Оставляет на каждую активную машину один последний ``gsm_waybill`` с
``source='imported'`` (статус ``exported``), удаляет остальные ПЛ,
все ``gsm_transaction`` и ``gsm_import_batch``.

Справочники (машины, водители, карты, маршруты, станции, настройки)
не трогает.

По умолчанию — dry-run (только отчёт). Запись: ``--apply``
(сначала sqlite backup API → ``*.bak-before-gsm-test-YYYYMMDD-HHMMSS``).

Пример:
  .venv/bin/python scripts/reset_gsm_to_anchors.py --db plita.db
  .venv/bin/python scripts/reset_gsm_to_anchors.py --db plita.db --apply
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Sequence

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from core.gsm.reset_to_anchors import ResetGsmError, format_plan, run_reset
from core.kp_db_common import DEFAULT_DB


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Сброс ГСМ к последним imported-якорям: удаляет прочие ПЛ, "
            "транзакции и батчи импорта. По умолчанию dry-run."
        )
    )
    parser.add_argument(
        "--db",
        type=Path,
        default=Path(DEFAULT_DB),
        help=f"Путь к SQLite (по умолчанию: {DEFAULT_DB})",
    )
    parser.add_argument(
        "--apply",
        action="store_true",
        help="Выполнить сброс (с бэкапом). Без флага — только отчёт.",
    )
    return parser.parse_args(list(argv) if argv is not None else None)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        result = run_reset(db_path=args.db.resolve(), apply=bool(args.apply))
        print(format_plan(result.plan, apply=bool(args.apply)))
        if not args.apply:
            print("  (ничего не записано; передайте --apply для выполнения)")
            return 0
        print(f"  бэкап: {result.backup_path}")
        plan = result.plan
        print(
            "  после сброса: "
            f"waybills={len(plan.anchors)} txs=0 batches=0 "
            f"routes={plan.routes_total} cards={plan.cards_total}"
        )
    except ResetGsmError as exc:
        print(f"Ошибка: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
