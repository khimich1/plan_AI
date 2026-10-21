"""Import a Т-13 workbook and serve the attendance grid."""

from __future__ import annotations

import calendar
import re
from collections import Counter
from datetime import date, datetime
from typing import Any

from app.repositories.gsm_repository import GsmRepository
from app.schemas.gsm import (
    TimesheetBatchOut,
    TimesheetGrid,
    TimesheetGridCell,
    TimesheetImportReport,
    TimesheetUnmatchedRow,
)
from core.gsm.timesheet import (
    TimesheetParseError,
    parse_personnel_number,
    parse_timesheet_xlsx,
)

_MONTH = re.compile(r"\d{4}-\d{2}")


class GsmTimesheetError(ValueError):
    """Domain error for timesheet import (maps to HTTP 400)."""

    def __init__(
        self,
        message: str,
        *,
        code: str = "gsm_timesheet_invalid",
        details: dict[str, Any] | None = None,
    ) -> None:
        super().__init__(message)
        self.code = code
        self.details = details or {}


class GsmTimesheetService:
    def __init__(self, *, repo: GsmRepository) -> None:
        self._repo = repo

    def import_xlsx(
        self,
        *,
        filename: str,
        content: bytes,
        month: str | date | None = None,
        uploaded_by: str | None = None,
    ) -> TimesheetImportReport:
        if not str(filename).lower().endswith(".xlsx"):
            raise GsmTimesheetError(
                "Нужен файл .xlsx формы Т-13.",
                code="gsm_timesheet_invalid",
            )
        month_date = _coerce_month(month)
        try:
            parsed = parse_timesheet_xlsx(
                content, filename=filename, month=month_date
            )
        except TimesheetParseError as exc:
            text = str(exc)
            code = (
                "gsm_timesheet_month_required"
                if "месяц" in text.lower()
                else "gsm_timesheet_invalid"
            )
            raise GsmTimesheetError(text, code=code) from exc

        period_from, period_to = parsed.period
        by_number = _active_drivers_by_number(self._repo.list_drivers(active_only=True))
        matched_rows = []
        unmatched: list[TimesheetUnmatchedRow] = [
            TimesheetUnmatchedRow(
                full_name=row.full_name,
                personnel_number=row.raw_personnel or None,
            )
            for row in parsed.unmatched
        ]
        for row in parsed.rows:
            driver = by_number.get(row.personnel_number)
            if driver is None:
                unmatched.append(
                    TimesheetUnmatchedRow(
                        full_name=row.full_name,
                        personnel_number=row.personnel_number,
                    )
                )
                continue
            matched_rows.append((int(driver["id"]), row))

        day_rows: list[dict[str, Any]] = []
        summary: Counter[str] = Counter()
        for driver_id, row in matched_rows:
            for day, code in row.codes:
                day_rows.append({"driver_id": driver_id, "date": day, "code": code})
                summary[code] += 1

        rows_total = len(parsed.rows) + len(parsed.unmatched)
        batch_id = self._repo.insert_timesheet_batch(
            filename=filename,
            period_from=period_from,
            period_to=period_to,
            rows_total=rows_total,
            matched=len(matched_rows),
            uploaded_by=uploaded_by,
            uploaded_at=datetime.now().isoformat(timespec="seconds"),
        )
        self._repo.replace_timesheet_days(
            period_from=period_from,
            period_to=period_to,
            batch_id=batch_id,
            rows=day_rows,
        )
        return TimesheetImportReport(
            period=period_from.strftime("%Y-%m"),
            rows_total=rows_total,
            matched=len(matched_rows),
            unmatched=unmatched,
            codes_summary=dict(summary),
        )

    def grid(self, period: str) -> TimesheetGrid:
        period_from, period_to, label = _month_bounds(period)
        cells = [
            TimesheetGridCell(
                driver_id=int(row["driver_id"]),
                full_name=str(row["full_name"]),
                personnel_number=row.get("personnel_number"),
                date=str(row["date"]),
                code=str(row["code"]),
            )
            for row in self._repo.list_timesheet_days(
                period_from=period_from, period_to=period_to
            )
        ]
        batches = [
            TimesheetBatchOut(
                id=int(row["id"]),
                filename=str(row["filename"]),
                period_from=str(row["period_from"]),
                period_to=str(row["period_to"]),
                rows_total=row.get("rows_total"),
                matched=row.get("matched"),
                uploaded_at=str(row["uploaded_at"]),
            )
            for row in self._repo.list_timesheet_batches()
        ]
        return TimesheetGrid(period=label, cells=cells, batches=batches)


def _active_drivers_by_number(drivers: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    found: dict[str, dict[str, Any]] = {}
    for driver in drivers:
        key = _personnel_key(driver.get("personnel_number"))
        if key is not None:
            found[key] = driver
    return found


def _personnel_key(value: object) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    if text.isdigit():
        return str(int(text))
    return parse_personnel_number(text)


def _coerce_month(month: str | date | None) -> date | None:
    if month is None or month == "":
        return None
    if isinstance(month, date):
        return month
    text = str(month).strip()
    if not _MONTH.fullmatch(text):
        raise GsmTimesheetError(
            "Месяц укажите как ГГГГ-ММ.",
            code="gsm_timesheet_invalid",
        )
    year_s, month_s = text.split("-")
    return date(int(year_s), int(month_s), 1)


def _month_bounds(period: str) -> tuple[date, date, str]:
    text = str(period).strip()
    if not _MONTH.fullmatch(text):
        raise GsmTimesheetError(
            "Период укажите как ГГГГ-ММ.",
            code="gsm_timesheet_invalid",
        )
    year_s, month_s = text.split("-")
    year_n, month_n = int(year_s), int(month_s)
    start = date(year_n, month_n, 1)
    end = date(year_n, month_n, calendar.monthrange(year_n, month_n)[1])
    return start, end, text
