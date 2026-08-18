"""Pure urgent-position helpers for production planning (no I/O, no app imports).

Aggregates per-plate deadlines from delivery batches and KP ``execution_terms``.
When a plate has no batch ``produce_by`` and ``execution_terms`` cannot be parsed,
falls back to ``now.date() + DEFAULT_EXECUTION_TERMS_DAYS`` (Phase 0 backlog
surface) with ``deadline_source="execution_terms"``. The +14 fallback is never
used for conflict comparison — only a real parsed KP deadline counts.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import Any, Mapping, Sequence

from core.execution_terms import (
    DEFAULT_EXECUTION_TERMS_DAYS,
    parse_execution_terms_to_datetime,
)

_CONFLICT_DAYS = 7


@dataclass(frozen=True, slots=True)
class UrgentPosition:
    plate_id: int
    kp_id: int
    plate_name: str
    qty_remaining: int
    deadline: date
    deadline_source: str  # "delivery_batch" | "execution_terms"
    deadline_details: list[dict[str, Any]]
    conflict: str | None  # "schedule_earlier" | "kp_earlier" | None


def _to_date(value: date | datetime | str) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    if "T" in text:
        text = text.split("T", 1)[0]
    return date.fromisoformat(text)


def _plate_id(plate: Mapping[str, Any]) -> int:
    raw = plate.get("plate_id", plate.get("id"))
    if raw is None:
        raise KeyError("plate map must contain plate_id or id")
    return int(raw)


def _batch_deadlines(
    batches: Sequence[Mapping[str, Any]],
) -> list[tuple[date, Mapping[str, Any]]]:
    result: list[tuple[date, Mapping[str, Any]]] = []
    for batch in batches:
        raw = batch.get("produce_by")
        if raw is None or raw == "":
            continue
        result.append((_to_date(raw), batch))
    return result


def _conflict(
    earliest_batch: date | None,
    kp_deadline: date | None,
) -> str | None:
    if earliest_batch is None or kp_deadline is None:
        return None
    if abs((earliest_batch - kp_deadline).days) <= _CONFLICT_DAYS:
        return None
    if earliest_batch < kp_deadline:
        return "schedule_earlier"
    return "kp_earlier"


def collect_urgent_positions(
    plates: Sequence[Mapping[str, Any]],
    batches_by_plate: Mapping[int, Sequence[Mapping[str, Any]]],
    kp_meta: Mapping[int, Mapping[str, Any]],
    deadline_until: date,
    *,
    now: datetime | None = None,
) -> list[UrgentPosition]:
    """Collect positions whose primary deadline is on or before ``deadline_until``.

    Primary deadline priority: earliest batch ``produce_by``, else parsed KP
    ``execution_terms``, else ``now.date() + DEFAULT_EXECUTION_TERMS_DAYS``.

    One ``UrgentPosition`` per plate. Sorted by deadline, then ``kp_id``, then
    ``plate_id``.
    """
    clock = now if now is not None else datetime.now()
    today = clock.date() if isinstance(clock, datetime) else clock

    urgent: list[UrgentPosition] = []
    for plate in plates:
        pid = _plate_id(plate)
        kp_id = int(plate["kp_id"])
        plate_name = str(plate.get("plate_name") or "")
        qty_remaining = int(plate.get("qty_remaining") or 0)

        batches = list(batches_by_plate.get(pid) or ())
        dated_batches = _batch_deadlines(batches)
        earliest_batch = (
            min(d for d, _ in dated_batches) if dated_batches else None
        )

        meta = kp_meta.get(kp_id) or {}
        terms_raw = str(meta.get("execution_terms") or "")
        parsed_terms = parse_execution_terms_to_datetime(terms_raw, now=clock)
        kp_deadline = parsed_terms.date() if parsed_terms is not None else None

        details: list[dict[str, Any]] = []
        for batch_date, batch in dated_batches:
            details.append(
                {
                    "type": "delivery_batch",
                    "batch_name": batch.get("batch_name"),
                    "deadline": batch_date.isoformat(),
                    "qty": batch.get("qty"),
                }
            )
        if kp_deadline is not None:
            details.append(
                {
                    "type": "execution_terms",
                    "deadline": kp_deadline.isoformat(),
                    "qty": qty_remaining,
                }
            )

        if earliest_batch is not None:
            deadline = earliest_batch
            deadline_source = "delivery_batch"
        elif kp_deadline is not None:
            deadline = kp_deadline
            deadline_source = "execution_terms"
        else:
            # Phase 0: unparseable terms + no batches → still surface in backlog
            deadline = today + timedelta(days=DEFAULT_EXECUTION_TERMS_DAYS)
            deadline_source = "execution_terms"
            details.append(
                {
                    "type": "execution_terms",
                    "deadline": deadline.isoformat(),
                    "qty": qty_remaining,
                }
            )

        if deadline > deadline_until:
            continue

        urgent.append(
            UrgentPosition(
                plate_id=pid,
                kp_id=kp_id,
                plate_name=plate_name,
                qty_remaining=qty_remaining,
                deadline=deadline,
                deadline_source=deadline_source,
                deadline_details=details,
                conflict=_conflict(earliest_batch, kp_deadline),
            )
        )

    urgent.sort(key=lambda p: (p.deadline, p.kp_id, p.plate_id))
    return urgent
