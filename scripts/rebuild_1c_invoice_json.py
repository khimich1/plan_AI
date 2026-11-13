#!/usr/bin/env python3
"""Восстановить JSON «КоммерческоеПредложение» для 1С без повторной отправки.

Тот же документ, что пишется в new_{kp_id}_*.json при export_invoice,
но без смены статуса и без файла в to_1c.

Пример на сервере:

  docker exec plan_zavod-backend-1 python /app/scripts/rebuild_1c_invoice_json.py 211 212

Или с хоста (если есть доступ к базе в томе):

  cd /opt/plan_zavod && .venv/bin/python scripts/rebuild_1c_invoice_json.py \\
    --db /var/lib/docker/volumes/plan_zavod_web_split_data/_data/plita.db 211 212
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app.repositories.kp_archive_repository import KpArchiveRepository
from app.services.archive_service import (
    ArchiveNotFoundError,
    ArchiveService,
    ArchiveValidationError,
    _warehouse_for_correction,
    _invoice_order_ids,
)
from core.invoice_export import INVOICE_WAREHOUSES, InvoiceAction


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Собрать JSON счёта для 1С по kp_id")
    parser.add_argument("kp_ids", nargs="+", type=int, help="Номера КП, например 211 212")
    parser.add_argument(
        "--db",
        type=Path,
        default=None,
        help="Путь к plita.db (по умолчанию из настроек приложения)",
    )
    parser.add_argument(
        "--warehouse",
        choices=list(INVOICE_WAREHOUSES),
        default=None,
        help="Склад; если не задан — из invoice_warehouse КП или «Склад Готовой Продукции»",
    )
    parser.add_argument(
        "--action",
        choices=("create", "update"),
        default="create",
        help="create — первая отправка (без УИД); update — delta с УИД из БД",
    )
    parser.add_argument(
        "--out-dir",
        type=Path,
        default=None,
        help="Папка для new_{id}.json; без флага — только stdout",
    )
    return parser.parse_args()


def _warehouse(raw: dict, requested: str | None) -> str:
    if requested:
        return requested
    stored = raw.get("invoice_warehouse")
    if stored and str(stored).strip() in INVOICE_WAREHOUSES:
        return str(stored).strip()
    return INVOICE_WAREHOUSES[0]


def _build_document(
    service: ArchiveService,
    kp_id: int,
    *,
    action: InvoiceAction,
    warehouse: str | None,
) -> dict:
    raw = service.repository.get_by_id(kp_id)
    if not raw:
        raise ArchiveNotFoundError(f"КП №{kp_id} не найдено")

    name = (
        _warehouse_for_correction(raw, warehouse)
        if action == "update"
        else _warehouse(raw, warehouse)
    )

    order_number: str | None = None
    uid_order: str | None = None
    uid_kp: str | None = None
    if action == "update":
        ids = _invoice_order_ids(raw)
        if ids is None:
            raise ArchiveValidationError(
                "Для update нет order_number_1c / uid_order_1c / uid_kp_1c"
            )
        order_number, uid_order, uid_kp = ids

    offer, contract = service._invoice_offer(raw)
    document, _digest = service._build_stamped_invoice(
        raw,
        offer,
        contract,
        action=action,
        warehouse=name,
        order_number=order_number,
        uid_order=uid_order,
        uid_kp=uid_kp,
    )
    return document


def main() -> int:
    args = _parse_args()
    repo = KpArchiveRepository(db_path=str(args.db)) if args.db else KpArchiveRepository()
    service = ArchiveService(repository=repo)
    exit_code = 0

    for kp_id in args.kp_ids:
        try:
            document = _build_document(
                service,
                kp_id,
                action=args.action,
                warehouse=args.warehouse,
            )
            payload = json.dumps(document, ensure_ascii=False, indent=2) + "\n"
            if args.out_dir is not None:
                args.out_dir.mkdir(parents=True, exist_ok=True)
                path = args.out_dir / f"new_{kp_id}_rebuilt.json"
                path.write_text(payload, encoding="utf-8")
                print(f"OK {kp_id} -> {path}", file=sys.stderr)
            print(f"===== КП {kp_id} =====")
            print(payload)
        except (ArchiveNotFoundError, ArchiveValidationError) as exc:
            exit_code = 1
            print(f"ОШИБКА КП {kp_id}: {exc}", file=sys.stderr)
        except Exception as exc:
            exit_code = 1
            print(f"ОШИБКА КП {kp_id}: {type(exc).__name__}: {exc}", file=sys.stderr)

    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
