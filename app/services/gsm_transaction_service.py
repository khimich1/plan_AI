"""Import fuel-card transaction .xls files into gsm_* tables."""

from __future__ import annotations

import sqlite3
from datetime import date, datetime
from typing import Any, Sequence

from app.repositories.gsm_repository import GsmRepository
from app.schemas.gsm import FileImportReport, TransactionImportReport, TransactionListResponse, TransactionOut
from core.gsm.transactions import ParsedTxFile, parse_transactions_content


class GsmTransactionError(Exception):
    """Domain error for transaction listing (invalid period)."""

    def __init__(self, message: str, *, code: str) -> None:
        super().__init__(message)
        self.code = code
        self.details: dict[str, object] = {}


class GsmTransactionService:
    """Multi-file transaction import with dedup and unmatched-card accept."""

    def __init__(self, *, repo: GsmRepository) -> None:
        self._repo = repo

    def import_files(
        self,
        files: Sequence[tuple[str, bytes]],
        *,
        uploaded_by: str | None = None,
    ) -> TransactionImportReport:
        file_reports: list[FileImportReport] = []
        total_inserted = 0
        total_duplicate = 0

        for filename, content in files:
            parsed = parse_transactions_content(content, filename=filename)
            report = self._import_parsed(parsed, uploaded_by=uploaded_by)
            file_reports.append(report)
            total_inserted += report.rows_inserted
            total_duplicate += report.rows_duplicate

        return TransactionImportReport(
            files=file_reports,
            rows_inserted=total_inserted,
            rows_duplicate=total_duplicate,
        )

    def _import_parsed(
        self,
        parsed: ParsedTxFile,
        *,
        uploaded_by: str | None,
    ) -> FileImportReport:
        period_from: str | None = None
        period_to: str | None = None
        if parsed.rows:
            timestamps = [row.ts for row in parsed.rows]
            period_from = min(timestamps).date().isoformat()
            period_to = max(timestamps).date().isoformat()

        batch_id = self._repo.create_import_batch(
            filename=parsed.filename,
            uploaded_at=datetime.now().isoformat(timespec="seconds"),
            uploaded_by=uploaded_by,
            period_from=period_from,
            period_to=period_to,
            rows_total=len(parsed.rows),
            sum_liters=parsed.sum_liters,
            sum_amount=parsed.sum_amount,
        )

        inserted = 0
        duplicate = 0
        unmatched: list[str] = []
        unmatched_seen: set[str] = set()
        card_cache: dict[str, int] = {}

        for row in parsed.rows:
            card_id = card_cache.get(row.card_number)
            if card_id is None:
                card_id, was_unmatched = self._resolve_card(row.card_number, row.ts)
                card_cache[row.card_number] = card_id
                if was_unmatched and row.card_number not in unmatched_seen:
                    unmatched_seen.add(row.card_number)
                    unmatched.append(row.card_number)

            station_id: int | None = None
            if row.raw_address:
                station_id = self._repo.get_or_create_station(
                    address=row.raw_address,
                    brand=row.brand or None,
                )

            ts_iso = row.ts.isoformat(sep=" ", timespec="seconds")
            qty_liters = None if row.service_type == "wash" else row.qty_liters
            try:
                self._repo.insert_transaction(
                    card_id=card_id,
                    ts=ts_iso,
                    service_type=row.service_type,
                    amount=row.amount,
                    raw_address=row.raw_address,
                    batch_id=batch_id,
                    qty_liters=qty_liters,
                    fuel_grade=row.fuel_grade,
                    station_id=station_id,
                )
                inserted += 1
            except sqlite3.IntegrityError:
                duplicate += 1

        return FileImportReport(
            filename=parsed.filename,
            rows_total=len(parsed.rows),
            rows_inserted=inserted,
            rows_duplicate=duplicate,
            sum_liters=parsed.sum_liters,
            sum_amount=parsed.sum_amount,
            footer_liters=parsed.footer_liters,
            footer_amount=parsed.footer_amount,
            warnings=list(parsed.warnings),
            unmatched_cards=unmatched,
        )

    def _resolve_card(self, card_number: str, ts: datetime) -> tuple[int, bool]:
        existing = self._repo.get_card_by_number(card_number)
        if existing is not None:
            return int(existing["id"]), False

        assigned_at = ts.date().isoformat() if isinstance(ts, datetime) else date.today().isoformat()
        card_id = self._repo.create_card(
            card_number=card_number,
            vehicle_id=None,
            assigned_at=assigned_at,
        )
        return card_id, True

    def list_transactions(
        self,
        *,
        period_from: date,
        period_to: date,
        vehicle_id: int | None = None,
        service_type: str | None = None,
    ) -> TransactionListResponse:
        if period_to < period_from:
            raise GsmTransactionError(
                "period_to must be >= period_from",
                code="gsm_invalid_period",
            )
        rows = self._repo.list_transactions(
            vehicle_id=vehicle_id,
            period_from=period_from,
            period_to=period_to,
            service_type=service_type,
        )
        items = [_transaction_out(row) for row in rows]
        sum_liters = round(sum((item.qty_liters or 0.0) for item in items), 2)
        sum_amount = round(sum(item.amount for item in items), 2)
        return TransactionListResponse(
            rows=items,
            total_count=len(items),
            sum_liters=sum_liters,
            sum_amount=sum_amount,
        )


def _transaction_out(row: dict[str, Any]) -> TransactionOut:
    qty = row.get("qty_liters")
    return TransactionOut(
        ts=str(row["ts"]),
        card_number=str(row["card_number"]),
        vehicle_id=row.get("vehicle_id"),
        service_type=str(row["service_type"]),
        fuel_grade=row.get("fuel_grade"),
        qty_liters=None if qty is None else round(float(qty), 2),
        amount=round(float(row["amount"]), 2),
        station_id=row.get("station_id"),
        address=row.get("raw_address"),
    )
