"""Seed a covering Т-13 so existing generate/export tests stay on code Я."""

from __future__ import annotations

from datetime import date, timedelta

from app.repositories.gsm_repository import GsmRepository


def seed_present_days(
    repo: GsmRepository,
    driver_id: int,
    *,
    period_from: date = date(2025, 1, 1),
    period_to: date = date(2026, 12, 31),
    code: str = "Я",
) -> int:
    """One batch plus a Я (or other code) on every day. Does not wipe other drivers."""
    batch_id = repo.insert_timesheet_batch(
        filename="seed от 31.12.2026.xlsx",
        period_from=period_from,
        period_to=period_to,
        rows_total=1,
        matched=1,
        uploaded_by="test",
        uploaded_at="2025-01-01T00:00:00",
    )
    rows: list[dict] = []
    day = period_from
    while day <= period_to:
        rows.append({"driver_id": driver_id, "date": day, "code": code})
        day += timedelta(days=1)
    repo.add_timesheet_days(batch_id=batch_id, rows=rows)
    return batch_id
