"""Pure helpers for GSM usage-report rows, tx attachment, and date labels."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta
from typing import Any, Sequence

from core.gsm.balance import burn_for_km
from core.gsm.blank import short_driver_name, vehicle_mark_label
from core.gsm.season import SeasonSwitch, norm_for

_MONTHS_NOMINATIVE = (
    "",
    "январь",
    "февраль",
    "март",
    "апрель",
    "май",
    "июнь",
    "июль",
    "август",
    "сентябрь",
    "октябрь",
    "ноябрь",
    "декабрь",
)

_MONTHS_GENITIVE = (
    "",
    "января",
    "февраля",
    "марта",
    "апреля",
    "мая",
    "июня",
    "июля",
    "августа",
    "сентября",
    "октября",
    "ноября",
    "декабря",
)

CONFIRMED_STATUSES = frozenset({"confirmed", "exported"})
KIT_STATUSES = frozenset({"draft", "confirmed", "exported"})


@dataclass(frozen=True, slots=True)
class UsageReportRow:
    seq: int
    day: date
    vehicle_mark: str
    plate: str
    driver_short: str
    fuel_grade: str
    fuel_start: float
    odometer_start: int
    odometer_end: int
    km: int
    norm_l_per_100: float
    burn_norm: float
    burn_fact: float
    received: float
    fuel_end: float
    note: str
    destination: str


@dataclass(frozen=True, slots=True)
class UsageMonthBlock:
    year: int
    month: int
    title: str  # «за  май  2026г.»
    rows: tuple[UsageReportRow, ...]
    fuel_start: float
    burn_norm: float
    burn_fact: float
    received: float
    fuel_end: float
    fuel_grade: str


def format_note_date(day: date) -> str:
    """«04 мая» style note column."""
    return f"{day.day:02d} {_MONTHS_GENITIVE[day.month]}"


def format_month_title(year: int, month: int) -> str:
    return f"за  {_MONTHS_NOMINATIVE[month]}  {year}г."


def format_approval_date(day: date) -> str:
    return f"{day.day} {_MONTHS_GENITIVE[day.month]} {day.year}г."


def report_filename(plate: str) -> str:
    """Zip entry: «Отчет по использованию ГСМ <госномер>.xls»."""
    plate_clean = " ".join((plate or "").split())
    return f"Отчет по использованию ГСМ {plate_clean}.xls"


def destination_from_route(route_items: Sequence[dict[str, Any]]) -> str:
    """Best-effort route destination (first non-Завод «to», else last «to»)."""
    if not route_items:
        return ""
    candidates: list[str] = []
    for item in route_items:
        to_addr = str(item.get("to") or item.get("to_addr") or item.get("addr_b") or "").strip()
        if to_addr:
            candidates.append(to_addr)
    for addr in candidates:
        low = addr.lower()
        if "завод" not in low and "азс" not in low:
            return addr
    return candidates[-1] if candidates else ""


def km_from_waybill(row: dict[str, Any], route_items: Sequence[dict[str, Any]]) -> int:
    odo_s = row.get("odometer_start")
    odo_e = row.get("odometer_end")
    if odo_s is not None and odo_e is not None:
        return max(0, int(odo_e) - int(odo_s))
    total = 0
    for item in route_items:
        try:
            total += int(item.get("km") or 0)
        except (TypeError, ValueError):
            continue
    return total


def attach_transactions_to_rows(
    row_dates: Sequence[date],
    transactions: Sequence[tuple[date, float]],
) -> list[float]:
    """Assign each tx qty to nearest PL row within ±1 calendar day.

    Tie-break: same date preferred, else earlier row date.
    Tx with no row in ±1 day are dropped (stage 2).
    """
    received = [0.0] * len(row_dates)
    if not row_dates:
        return received

    for tx_day, qty in transactions:
        best_i: int | None = None
        best_key: tuple[int, int, int] | None = None  # (dist, same_day_penalty, index)
        for i, row_day in enumerate(row_dates):
            dist = abs((tx_day - row_day).days)
            if dist > 1:
                continue
            # Prefer same date (penalty 0), else earlier row (smaller index)
            same_penalty = 0 if dist == 0 else 1
            key = (dist, same_penalty, i)
            if best_key is None or key < best_key:
                best_key = key
                best_i = i
        if best_i is not None:
            received[best_i] = round(received[best_i] + float(qty), 2)
    return received


def build_month_block(
    *,
    year: int,
    month: int,
    waybills: Sequence[dict[str, Any]],
    vehicle: dict[str, Any],
    drivers_by_id: dict[int, dict[str, Any]],
    route_by_wb_id: dict[int, list[dict[str, Any]]],
    received_by_date: dict[date, float],
    season_switches: Sequence[SeasonSwitch],
    fuel_grade: str = "АИ-95",
) -> UsageMonthBlock | None:
    """Build one month block from confirmed/exported waybills (1 PL = 1 row)."""
    if not waybills:
        return None

    mark = vehicle_mark_label(str(vehicle.get("name") or ""))
    plate = str(vehicle.get("plate_number") or "")
    rows: list[UsageReportRow] = []

    for seq, wb in enumerate(waybills, start=1):
        day = _as_date(wb["date"])
        route = route_by_wb_id.get(int(wb["id"]), [])
        km = km_from_waybill(wb, route)
        norm = norm_for(
            day,
            norm_summer=float(vehicle["norm_summer"]),
            norm_winter=float(vehicle["norm_winter"]),
            switches=season_switches,
        )
        burn = burn_for_km(km, norm)
        driver = drivers_by_id.get(int(wb["driver_id"]))
        driver_short = short_driver_name(str(driver["full_name"])) if driver else ""
        fuel_start = float(wb.get("fuel_start") or 0.0)
        fuel_end = float(wb.get("fuel_end") or 0.0)
        odo_start = int(wb.get("odometer_start") or 0)
        odo_end = int(wb.get("odometer_end") or (odo_start + km))
        received = float(received_by_date.get(day, 0.0))
        rows.append(
            UsageReportRow(
                seq=seq,
                day=day,
                vehicle_mark=mark,
                plate=plate,
                driver_short=driver_short,
                fuel_grade=fuel_grade,
                fuel_start=fuel_start,
                odometer_start=odo_start,
                odometer_end=odo_end,
                km=km,
                norm_l_per_100=float(norm),
                burn_norm=burn,
                burn_fact=burn,
                received=received,
                fuel_end=fuel_end,
                note=format_note_date(day),
                destination=destination_from_route(route),
            )
        )

    first = rows[0]
    last = rows[-1]
    return UsageMonthBlock(
        year=year,
        month=month,
        title=format_month_title(year, month),
        rows=tuple(rows),
        fuel_start=first.fuel_start,
        burn_norm=round(sum(r.burn_norm for r in rows), 2),
        burn_fact=round(sum(r.burn_fact for r in rows), 2),
        received=round(sum(r.received for r in rows), 2),
        fuel_end=last.fuel_end,
        fuel_grade=fuel_grade,
    )


def group_waybills_by_month(
    waybills: Sequence[dict[str, Any]],
) -> list[tuple[int, int, list[dict[str, Any]]]]:
    """Group sorted waybills into (year, month, rows) chronologically."""
    groups: dict[tuple[int, int], list[dict[str, Any]]] = {}
    order: list[tuple[int, int]] = []
    for wb in waybills:
        day = _as_date(wb["date"])
        key = (day.year, day.month)
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(wb)
    return [(y, m, groups[(y, m)]) for y, m in order]


def _as_date(value: Any) -> date:
    if isinstance(value, date):
        return value
    return date.fromisoformat(str(value)[:10])


def tx_day_qty(tx: dict[str, Any]) -> tuple[date, float] | None:
    """Extract (day, qty) for fuel txs; skip non-fuel / missing qty."""
    if str(tx.get("service_type") or "") != "fuel":
        return None
    qty = tx.get("qty_liters")
    if qty is None:
        return None
    ts = tx.get("ts")
    if not ts:
        return None
    return date.fromisoformat(str(ts)[:10]), float(qty)


def expand_tx_window(period_from: date, period_to: date) -> tuple[date, date]:
    """Tx fetch window = period ±1 day for nearest-row attachment."""
    return period_from - timedelta(days=1), period_to + timedelta(days=1)
