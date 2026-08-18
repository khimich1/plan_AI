"""Import fuel-card transaction .xls files into gsm_* tables."""

from __future__ import annotations

import sqlite3
from datetime import date, datetime
from typing import Sequence

from app.repositories.gsm_repository import GsmRepository
from app.schemas.gsm import FileImportReport, TransactionImportReport
from core.gsm.transactions import ParsedTxFile, parse_transactions_content


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
            try:
                self._repo.insert_transaction(
                    card_id=card_id,
                    ts=ts_iso,
                    service_type=row.service_type,
                    amount=row.amount,
                    raw_address=row.raw_address,
                    batch_id=batch_id,
                    qty_liters=row.qty_liters,
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
