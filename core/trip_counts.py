"""Принятое число рейсов трёх простых котлов КП.

Формулы рейсов сюда не входят: вызывающий передаёт уже посчитанные
массу и целое формулы. Здесь только подмена перед умножением на тариф.
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from typing import Any, Mapping

TRIP_BUCKETS = ("plate", "pile", "fbs_lm")
MASS_RESET_EPSILON_KG = 0.01


@dataclass(frozen=True)
class TripBucketInput:
    present: bool
    ready: bool
    cargo_kg: float
    formula: int


@dataclass(frozen=True)
class ResolvedTrips:
    accepted: dict[str, int]
    formula: dict[str, int]
    notices: dict[str, dict[str, int]]
    stored: dict[str, Any]

    def notice_for(self, bucket: str) -> dict[str, int] | None:
        return self.notices.get(bucket)


def _as_int(value: Any) -> int | None:
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return None


def _as_mass(value: Any) -> float | None:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return None
    return float(value)


def _parse_record(raw: Any) -> dict[str, Any] | None:
    if not isinstance(raw, Mapping):
        return None
    accepted = _as_int(raw.get("accepted"))
    basis = _as_mass(raw.get("basis_kg"))
    formula = _as_int(raw.get("formula"))
    if accepted is None or accepted < 0 or basis is None or formula is None or formula < 0:
        return None
    return {"accepted": accepted, "basis_kg": basis, "formula": formula}


def _parse_notice(raw: Any) -> dict[str, Any] | None:
    if not isinstance(raw, Mapping):
        return None
    bucket = str(raw.get("bucket") or "")
    previous = _as_int(raw.get("previous"))
    nxt = _as_int(raw.get("next"))
    if bucket not in TRIP_BUCKETS or previous is None or nxt is None:
        return None
    return {"bucket": bucket, "previous": previous, "next": nxt}


def coerce_trip_counts(raw: Any) -> dict[str, Any]:
    """Пустой объект, если JSON нет или он не читается."""
    payload = raw
    if isinstance(raw, str):
        text = raw.strip()
        if not text:
            payload = {}
        else:
            try:
                payload = json.loads(text)
            except json.JSONDecodeError:
                payload = {}
    if not isinstance(payload, Mapping):
        payload = {}
    stored: dict[str, Any] = {"notices": []}
    for bucket in TRIP_BUCKETS:
        record = _parse_record(payload.get(bucket))
        if record is not None:
            stored[bucket] = record
    notices_raw = payload.get("notices")
    if isinstance(notices_raw, list):
        stored["notices"] = [
            notice for item in notices_raw if (notice := _parse_notice(item)) is not None
        ]
    return stored


def dumps_trip_counts(stored: Any) -> str:
    return json.dumps(coerce_trip_counts(stored), ensure_ascii=False, separators=(",", ":"))


def dismiss_trip_notices(stored: Any, buckets: list[str] | tuple[str, ...]) -> dict[str, Any]:
    """Убирает строки напоминания. Принятые числа не меняет."""
    parsed = coerce_trip_counts(stored)
    drop = {bucket for bucket in buckets if bucket in TRIP_BUCKETS}
    parsed["notices"] = [row for row in parsed["notices"] if row["bucket"] not in drop]
    return parsed


def _override_value(overrides: Mapping[str, int] | None, bucket: str) -> int | None:
    if not overrides or bucket not in overrides:
        return None
    value = _as_int(overrides.get(bucket))
    if value is None or value < 0:
        return None
    return value


def _replace_notice(
    notices: list[dict[str, Any]],
    bucket: str,
    previous: int,
    nxt: int,
) -> list[dict[str, Any]]:
    kept = [row for row in notices if row["bucket"] != bucket]
    kept.append({"bucket": bucket, "previous": previous, "next": nxt})
    return kept


def resolve_trip_counts(
    stored: Any,
    *,
    plate: TripBucketInput,
    pile: TripBucketInput,
    fbs_lm: TripBucketInput,
    overrides: Mapping[str, int] | None = None,
) -> ResolvedTrips:
    parsed = coerce_trip_counts(stored)
    records = {bucket: parsed[bucket] for bucket in TRIP_BUCKETS if bucket in parsed}
    notices: list[dict[str, Any]] = list(parsed["notices"])
    inputs = {"plate": plate, "pile": pile, "fbs_lm": fbs_lm}
    accepted: dict[str, int] = {}
    formula: dict[str, int] = {}

    for bucket, spec in inputs.items():
        formula[bucket] = int(spec.formula) if spec.present else 0
        if not spec.present:
            accepted[bucket] = 0
            continue
        if not spec.ready:
            records.pop(bucket, None)
            notices = [row for row in notices if row["bucket"] != bucket]
            accepted[bucket] = 0
            continue

        explicit = _override_value(overrides, bucket)
        record = records.get(bucket)
        if record is None:
            chosen = int(spec.formula) if explicit is None else explicit
            records[bucket] = {
                "accepted": chosen,
                "basis_kg": float(spec.cargo_kg),
                "formula": int(spec.formula),
            }
            accepted[bucket] = chosen
            continue

        basis = float(record["basis_kg"])
        if abs(float(spec.cargo_kg) - basis) > MASS_RESET_EPSILON_KG + 1e-6:
            previous = int(record["accepted"])
            chosen = int(spec.formula)
            records[bucket] = {
                "accepted": chosen,
                "basis_kg": float(spec.cargo_kg),
                "formula": int(spec.formula),
            }
            if previous != chosen:
                notices = _replace_notice(notices, bucket, previous, chosen)
            accepted[bucket] = chosen
            continue

        chosen = int(record["accepted"])
        if explicit is not None:
            chosen = explicit
            records[bucket] = {
                "accepted": chosen,
                "basis_kg": basis,
                "formula": int(record["formula"]),
            }
        accepted[bucket] = chosen

    notice_map: dict[str, dict[str, int]] = {}
    for row in notices:
        bucket = row["bucket"]
        spec = inputs.get(bucket)
        if spec is None or not spec.present or not spec.ready:
            continue
        notice_map[bucket] = {"previous": int(row["previous"]), "next": int(row["next"])}

    stored_out: dict[str, Any] = {"notices": notices}
    for bucket in TRIP_BUCKETS:
        if bucket in records:
            stored_out[bucket] = records[bucket]
    return ResolvedTrips(
        accepted=accepted,
        formula=formula,
        notices=notice_map,
        stored=stored_out,
    )
