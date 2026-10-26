"""First clustered fill strictly after a period, one fuel day at a time."""

from __future__ import annotations

from datetime import date, datetime
from typing import Any

from app.repositories.gsm_repository import GsmRepository
from core.gsm.generator import _first_fill_liters, _station_ids
from core.gsm.models import Transaction

_SEARCH_DAYS = 48


def _row_to_transaction(row: dict[str, Any]) -> Transaction:
    ts_raw = row["ts"]
    ts = ts_raw if isinstance(ts_raw, datetime) else datetime.fromisoformat(str(ts_raw))
    qty = row.get("qty_liters")
    station_id = row.get("station_id")
    return Transaction(
        card_id=int(row["card_id"]),
        ts=ts,
        service_type=str(row.get("service_type") or "fuel"),
        qty_liters=None if qty is None else float(qty),
        amount=float(row.get("amount") or 0.0),
        station_id=None if station_id is None else int(station_id),
        raw_address=str(row.get("raw_address") or ""),
    )


def lookup_boundary_fill(
    repo: GsmRepository,
    *,
    vehicle_id: int,
    after: date,
) -> tuple[float, tuple[int, ...]] | None:
    """(liters, station_ids) of the first positive clustered fill after ``after``.

    Walks one fuel day at a time. Returns None when no later positive fill exists.
    """
    cursor = after
    for _ in range(_SEARCH_DAYS):
        grouped = repo.next_fuel_day_rows(after=cursor, vehicle_id=vehicle_id)
        rows = grouped.get(int(vehicle_id))
        if not rows:
            return None
        txs = tuple(_row_to_transaction(row) for row in rows)
        liters = round(_first_fill_liters(txs), 2)
        if liters > 0:
            return liters, _station_ids(txs)
        cursor = min(tx.ts.date() for tx in txs)
    return None


def attach_next_first_fills(
    repo: GsmRepository,
    rows: list[dict[str, Any]],
    *,
    after: date,
) -> None:
    """Set ``next_first_fill_liters`` on overview rows. None when no later fill."""
    for row in rows:
        found = lookup_boundary_fill(
            repo, vehicle_id=int(row["vehicle_id"]), after=after
        )
        row["next_first_fill_liters"] = None if found is None else found[0]
