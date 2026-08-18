"""SQL persistence for GSM module tables (vehicles, cards, transactions, waybills)."""

from __future__ import annotations

import sqlite3
from datetime import date, datetime
from typing import Any, Union

from core.kp_db_common import _connect
from core.kp_db_schema import ensure_schema

DateLike = Union[date, str]

_DRIVER_UPDATE_FIELDS = frozenset(
    {
        "full_name",
        "license_number",
        "license_issued_at",
        "personnel_number",
        "snils",
        "is_active",
    }
)
_VEHICLE_UPDATE_FIELDS = frozenset(
    {
        "name",
        "plate_number",
        "tank_volume_liters",
        "norm_summer",
        "norm_winter",
        "primary_driver_id",
        "is_active",
    }
)
_CARD_UPDATE_FIELDS = frozenset({"vehicle_id", "assigned_at", "archived_at"})
_STATION_UPDATE_FIELDS = frozenset(
    {"address", "brand", "lat", "lon", "geocode_source"}
)


class GsmRepository:
    """CRUD for ``gsm_*`` tables; cards are archived, never deleted."""

    def __init__(self, *, db_path: str) -> None:
        self.db_path = db_path

    def _connect(self) -> sqlite3.Connection:
        ensure_schema(self.db_path)
        conn = _connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    @staticmethod
    def _to_iso_date(value: DateLike) -> str:
        if isinstance(value, date):
            return value.isoformat()
        return date.fromisoformat(str(value)).isoformat()

    @staticmethod
    def _row_to_dict(row: sqlite3.Row | None) -> dict[str, Any] | None:
        if row is None:
            return None
        return dict(row)

    # ------------------------------------------------------------------
    # Drivers
    # ------------------------------------------------------------------

    def create_driver(
        self,
        *,
        full_name: str,
        license_number: str,
        license_issued_at: str | None = None,
        personnel_number: str | None = None,
        snils: str | None = None,
        is_active: int = 1,
    ) -> int:
        with self._connect() as conn:
            cur = conn.execute(
                """
                INSERT INTO gsm_driver (
                    full_name, license_number, license_issued_at,
                    personnel_number, snils, is_active
                ) VALUES (?, ?, ?, ?, ?, ?)
                """,
                (
                    full_name,
                    license_number,
                    license_issued_at,
                    personnel_number,
                    snils,
                    int(is_active),
                ),
            )
            conn.commit()
            return int(cur.lastrowid)

    def get_driver(self, driver_id: int) -> dict[str, Any] | None:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM gsm_driver WHERE id = ?",
                (driver_id,),
            ).fetchone()
        return self._row_to_dict(row)

    def list_drivers(self, *, active_only: bool = True) -> list[dict[str, Any]]:
        with self._connect() as conn:
            if active_only:
                rows = conn.execute(
                    "SELECT * FROM gsm_driver WHERE is_active = 1 ORDER BY id"
                ).fetchall()
            else:
                rows = conn.execute(
                    "SELECT * FROM gsm_driver ORDER BY id"
                ).fetchall()
        return [dict(row) for row in rows]

    def update_driver(self, driver_id: int, **fields: Any) -> None:
        unknown = set(fields) - _DRIVER_UPDATE_FIELDS
        if unknown:
            raise ValueError(f"Unknown driver fields: {sorted(unknown)}")
        if not fields:
            return
        assignments = ", ".join(f"{col} = ?" for col in fields)
        values = list(fields.values()) + [driver_id]
        with self._connect() as conn:
            conn.execute(
                f"UPDATE gsm_driver SET {assignments} WHERE id = ?",
                values,
            )
            conn.commit()

    # ------------------------------------------------------------------
    # Vehicles
    # ------------------------------------------------------------------

    def create_vehicle(
        self,
        *,
        name: str,
        plate_number: str,
        tank_volume_liters: float,
        norm_summer: float,
        norm_winter: float,
        primary_driver_id: int | None = None,
        is_active: int = 1,
    ) -> int:
        with self._connect() as conn:
            cur = conn.execute(
                """
                INSERT INTO gsm_vehicle (
                    name, plate_number, tank_volume_liters,
                    norm_summer, norm_winter, primary_driver_id, is_active
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    name,
                    plate_number,
                    float(tank_volume_liters),
                    float(norm_summer),
                    float(norm_winter),
                    primary_driver_id,
                    int(is_active),
                ),
            )
            conn.commit()
            return int(cur.lastrowid)

    def get_vehicle(self, vehicle_id: int) -> dict[str, Any] | None:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM gsm_vehicle WHERE id = ?",
                (vehicle_id,),
            ).fetchone()
        return self._row_to_dict(row)

    def list_vehicles(self, *, active_only: bool = True) -> list[dict[str, Any]]:
        with self._connect() as conn:
            if active_only:
                rows = conn.execute(
                    "SELECT * FROM gsm_vehicle WHERE is_active = 1 ORDER BY id"
                ).fetchall()
            else:
                rows = conn.execute(
                    "SELECT * FROM gsm_vehicle ORDER BY id"
                ).fetchall()
        return [dict(row) for row in rows]

    def update_vehicle(self, vehicle_id: int, **fields: Any) -> None:
        unknown = set(fields) - _VEHICLE_UPDATE_FIELDS
        if unknown:
            raise ValueError(f"Unknown vehicle fields: {sorted(unknown)}")
        if not fields:
            return
        assignments = ", ".join(f"{col} = ?" for col in fields)
        values = list(fields.values()) + [vehicle_id]
        with self._connect() as conn:
            conn.execute(
                f"UPDATE gsm_vehicle SET {assignments} WHERE id = ?",
                values,
            )
            conn.commit()

    # ------------------------------------------------------------------
    # Stations
    # ------------------------------------------------------------------

    def create_station(
        self,
        *,
        address: str,
        brand: str | None = None,
        lat: float | None = None,
        lon: float | None = None,
        geocode_source: str | None = None,
    ) -> int:
        with self._connect() as conn:
            cur = conn.execute(
                """
                INSERT INTO gsm_station (address, brand, lat, lon, geocode_source)
                VALUES (?, ?, ?, ?, ?)
                """,
                (address, brand, lat, lon, geocode_source),
            )
            conn.commit()
            return int(cur.lastrowid)

    def get_station(self, station_id: int) -> dict[str, Any] | None:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM gsm_station WHERE id = ?",
                (station_id,),
            ).fetchone()
        return self._row_to_dict(row)

    def get_station_by_address(self, address: str) -> dict[str, Any] | None:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM gsm_station WHERE address = ?",
                (address,),
            ).fetchone()
        return self._row_to_dict(row)

    def get_or_create_station(
        self,
        *,
        address: str,
        brand: str | None = None,
    ) -> int:
        existing = self.get_station_by_address(address)
        if existing is not None:
            return int(existing["id"])
        return self.create_station(address=address, brand=brand)

    def list_stations(self) -> list[dict[str, Any]]:
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT * FROM gsm_station ORDER BY id"
            ).fetchall()
        return [dict(row) for row in rows]

    def update_station(self, station_id: int, **fields: Any) -> None:
        unknown = set(fields) - _STATION_UPDATE_FIELDS
        if unknown:
            raise ValueError(f"Unknown station fields: {sorted(unknown)}")
        if not fields:
            return
        assignments = ", ".join(f"{col} = ?" for col in fields)
        values = list(fields.values()) + [station_id]
        with self._connect() as conn:
            conn.execute(
                f"UPDATE gsm_station SET {assignments} WHERE id = ?",
                values,
            )
            conn.commit()

    # ------------------------------------------------------------------
    # Fuel cards
    # ------------------------------------------------------------------

    def create_card(
        self,
        *,
        card_number: str,
        vehicle_id: int | None,
        assigned_at: str,
    ) -> int:
        with self._connect() as conn:
            cur = conn.execute(
                """
                INSERT INTO gsm_fuel_card (card_number, vehicle_id, assigned_at)
                VALUES (?, ?, ?)
                """,
                (card_number, vehicle_id, assigned_at),
            )
            conn.commit()
            return int(cur.lastrowid)

    def get_card(self, card_id: int) -> dict[str, Any] | None:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM gsm_fuel_card WHERE id = ?",
                (card_id,),
            ).fetchone()
        return self._row_to_dict(row)

    def get_card_by_number(self, card_number: str) -> dict[str, Any] | None:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM gsm_fuel_card WHERE card_number = ?",
                (card_number,),
            ).fetchone()
        return self._row_to_dict(row)

    def list_cards(self, *, include_archived: bool = False) -> list[dict[str, Any]]:
        with self._connect() as conn:
            if include_archived:
                rows = conn.execute(
                    "SELECT * FROM gsm_fuel_card ORDER BY id"
                ).fetchall()
            else:
                rows = conn.execute(
                    "SELECT * FROM gsm_fuel_card "
                    "WHERE archived_at IS NULL ORDER BY id"
                ).fetchall()
        return [dict(row) for row in rows]

    def archive_card(self, card_id: int) -> None:
        archived_at = datetime.now().isoformat(timespec="seconds")
        with self._connect() as conn:
            conn.execute(
                "UPDATE gsm_fuel_card SET archived_at = ? WHERE id = ?",
                (archived_at, card_id),
            )
            conn.commit()

    def update_card(self, card_id: int, **fields: Any) -> None:
        unknown = set(fields) - _CARD_UPDATE_FIELDS
        if unknown:
            raise ValueError(f"Unknown card fields: {sorted(unknown)}")
        if not fields:
            return
        assignments = ", ".join(f"{col} = ?" for col in fields)
        values = list(fields.values()) + [card_id]
        with self._connect() as conn:
            conn.execute(
                f"UPDATE gsm_fuel_card SET {assignments} WHERE id = ?",
                values,
            )
            conn.commit()

    # ------------------------------------------------------------------
    # Import batches + transactions
    # ------------------------------------------------------------------

    def create_import_batch(
        self,
        *,
        filename: str,
        uploaded_at: str,
        uploaded_by: str | None = None,
        period_from: str | None = None,
        period_to: str | None = None,
        rows_total: int | None = None,
        sum_liters: float | None = None,
        sum_amount: float | None = None,
    ) -> int:
        with self._connect() as conn:
            cur = conn.execute(
                """
                INSERT INTO gsm_import_batch (
                    filename, period_from, period_to, rows_total,
                    sum_liters, sum_amount, uploaded_by, uploaded_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    filename,
                    period_from,
                    period_to,
                    rows_total,
                    sum_liters,
                    sum_amount,
                    uploaded_by,
                    uploaded_at,
                ),
            )
            conn.commit()
            return int(cur.lastrowid)

    def insert_transaction(
        self,
        *,
        card_id: int,
        ts: str,
        service_type: str,
        amount: float,
        raw_address: str,
        batch_id: int,
        qty_liters: float | None = None,
        fuel_grade: str | None = None,
        station_id: int | None = None,
    ) -> int:
        with self._connect() as conn:
            cur = conn.execute(
                """
                INSERT INTO gsm_transaction (
                    card_id, ts, service_type, fuel_grade, qty_liters,
                    amount, station_id, raw_address, batch_id
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    card_id,
                    ts,
                    service_type,
                    fuel_grade,
                    qty_liters,
                    float(amount),
                    station_id,
                    raw_address,
                    batch_id,
                ),
            )
            conn.commit()
            return int(cur.lastrowid)

    def list_transactions(
        self,
        *,
        vehicle_id: int,
        period_from: DateLike,
        period_to: DateLike,
    ) -> list[dict[str, Any]]:
        from_iso = self._to_iso_date(period_from)
        to_iso = self._to_iso_date(period_to)
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT t.*
                FROM gsm_transaction t
                JOIN gsm_fuel_card c ON c.id = t.card_id
                WHERE c.vehicle_id = ?
                  AND substr(t.ts, 1, 10) >= ?
                  AND substr(t.ts, 1, 10) <= ?
                ORDER BY t.ts, t.id
                """,
                (vehicle_id, from_iso, to_iso),
            ).fetchall()
        return [dict(row) for row in rows]

    # ------------------------------------------------------------------
    # Routes
    # ------------------------------------------------------------------

    def create_route(
        self,
        *,
        vehicle_id: int,
        addr_a: str,
        addr_b: str,
        km: int,
        frequency: int = 1,
        typical_station_ids: str | None = None,
    ) -> int:
        with self._connect() as conn:
            cur = conn.execute(
                """
                INSERT INTO gsm_route (
                    vehicle_id, addr_a, addr_b, km, frequency, typical_station_ids
                ) VALUES (?, ?, ?, ?, ?, ?)
                """,
                (
                    vehicle_id,
                    addr_a,
                    addr_b,
                    int(km),
                    int(frequency),
                    typical_station_ids,
                ),
            )
            conn.commit()
            return int(cur.lastrowid)

    def list_routes(
        self,
        *,
        vehicle_id: int | None = None,
    ) -> list[dict[str, Any]]:
        with self._connect() as conn:
            if vehicle_id is None:
                rows = conn.execute(
                    "SELECT * FROM gsm_route ORDER BY id"
                ).fetchall()
            else:
                rows = conn.execute(
                    "SELECT * FROM gsm_route WHERE vehicle_id = ? ORDER BY id",
                    (vehicle_id,),
                ).fetchall()
        return [dict(row) for row in rows]

    # ------------------------------------------------------------------
    # Waybills
    # ------------------------------------------------------------------

    def upsert_waybill(
        self,
        *,
        vehicle_id: int,
        date: DateLike,
        driver_id: int,
        status: str = "draft",
        source: str = "auto",
        odometer_start: int | None = None,
        odometer_end: int | None = None,
        fuel_start: float | None = None,
        fuel_issued: float | None = None,
        fuel_end: float | None = None,
        route_json: str = "[]",
        warnings_json: str | None = None,
    ) -> int:
        day_iso = self._to_iso_date(date)
        with self._connect() as conn:
            cur = conn.execute(
                """
                INSERT INTO gsm_waybill (
                    vehicle_id, date, driver_id, status, source,
                    odometer_start, odometer_end,
                    fuel_start, fuel_issued, fuel_end,
                    route_json, warnings_json
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(vehicle_id, date) DO UPDATE SET
                    driver_id = excluded.driver_id,
                    status = excluded.status,
                    source = excluded.source,
                    odometer_start = excluded.odometer_start,
                    odometer_end = excluded.odometer_end,
                    fuel_start = excluded.fuel_start,
                    fuel_issued = excluded.fuel_issued,
                    fuel_end = excluded.fuel_end,
                    route_json = excluded.route_json,
                    warnings_json = excluded.warnings_json
                """,
                (
                    vehicle_id,
                    day_iso,
                    driver_id,
                    status,
                    source,
                    odometer_start,
                    odometer_end,
                    fuel_start,
                    fuel_issued,
                    fuel_end,
                    route_json,
                    warnings_json,
                ),
            )
            conn.commit()
            if cur.lastrowid:
                return int(cur.lastrowid)
            row = conn.execute(
                "SELECT id FROM gsm_waybill WHERE vehicle_id = ? AND date = ?",
                (vehicle_id, day_iso),
            ).fetchone()
            return int(row["id"])

    def get_waybill(
        self,
        vehicle_id: int,
        day: DateLike,
    ) -> dict[str, Any] | None:
        day_iso = self._to_iso_date(day)
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM gsm_waybill WHERE vehicle_id = ? AND date = ?",
                (vehicle_id, day_iso),
            ).fetchone()
        return self._row_to_dict(row)

    def get_waybill_by_id(self, waybill_id: int) -> dict[str, Any] | None:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM gsm_waybill WHERE id = ?",
                (waybill_id,),
            ).fetchone()
        return self._row_to_dict(row)

    def list_waybills(
        self,
        *,
        vehicle_id: int,
        period_from: DateLike,
        period_to: DateLike,
    ) -> list[dict[str, Any]]:
        from_iso = self._to_iso_date(period_from)
        to_iso = self._to_iso_date(period_to)
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT * FROM gsm_waybill
                WHERE vehicle_id = ?
                  AND date >= ?
                  AND date <= ?
                ORDER BY date, id
                """,
                (vehicle_id, from_iso, to_iso),
            ).fetchall()
        return [dict(row) for row in rows]

    def list_waybills_after(
        self,
        *,
        vehicle_id: int,
        after_date: DateLike,
    ) -> list[dict[str, Any]]:
        """Waybills strictly after ``after_date``, ascending by date."""
        after_iso = self._to_iso_date(after_date)
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT * FROM gsm_waybill
                WHERE vehicle_id = ?
                  AND date > ?
                ORDER BY date, id
                """,
                (vehicle_id, after_iso),
            ).fetchall()
        return [dict(row) for row in rows]

    def get_last_waybill(
        self,
        vehicle_id: int,
        *,
        before: DateLike,
    ) -> dict[str, Any] | None:
        """Latest waybill of any status strictly before ``before``."""
        before_iso = self._to_iso_date(before)
        with self._connect() as conn:
            row = conn.execute(
                """
                SELECT * FROM gsm_waybill
                WHERE vehicle_id = ?
                  AND date < ?
                ORDER BY date DESC, id DESC
                LIMIT 1
                """,
                (vehicle_id, before_iso),
            ).fetchone()
        return self._row_to_dict(row)

    def get_last_confirmed_waybill(
        self,
        vehicle_id: int,
        *,
        before: DateLike | None = None,
    ) -> dict[str, Any] | None:
        """Latest confirmed/exported waybill for vehicle, optionally before a date."""
        with self._connect() as conn:
            if before is None:
                row = conn.execute(
                    """
                    SELECT * FROM gsm_waybill
                    WHERE vehicle_id = ?
                      AND status IN ('confirmed', 'exported')
                    ORDER BY date DESC, id DESC
                    LIMIT 1
                    """,
                    (vehicle_id,),
                ).fetchone()
            else:
                before_iso = self._to_iso_date(before)
                row = conn.execute(
                    """
                    SELECT * FROM gsm_waybill
                    WHERE vehicle_id = ?
                      AND status IN ('confirmed', 'exported')
                      AND date < ?
                    ORDER BY date DESC, id DESC
                    LIMIT 1
                    """,
                    (vehicle_id, before_iso),
                ).fetchone()
        return self._row_to_dict(row)

    def delete_waybills_in_period(
        self,
        *,
        vehicle_id: int,
        period_from: DateLike,
        period_to: DateLike,
        statuses: tuple[str, ...] | None = None,
    ) -> int:
        """Delete waybills in period; optionally filter by status. Returns deleted count."""
        from_iso = self._to_iso_date(period_from)
        to_iso = self._to_iso_date(period_to)
        with self._connect() as conn:
            if statuses is None:
                cur = conn.execute(
                    """
                    DELETE FROM gsm_waybill
                    WHERE vehicle_id = ? AND date >= ? AND date <= ?
                    """,
                    (vehicle_id, from_iso, to_iso),
                )
            else:
                placeholders = ",".join("?" for _ in statuses)
                cur = conn.execute(
                    f"""
                    DELETE FROM gsm_waybill
                    WHERE vehicle_id = ? AND date >= ? AND date <= ?
                      AND status IN ({placeholders})
                    """,
                    (vehicle_id, from_iso, to_iso, *statuses),
                )
            conn.commit()
            return int(cur.rowcount)

    # ------------------------------------------------------------------
    # Settings
    # ------------------------------------------------------------------

    def get_setting(self, key: str) -> str | None:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT value FROM gsm_setting WHERE key = ?",
                (key,),
            ).fetchone()
        if row is None:
            return None
        return str(row["value"])

    def set_setting(self, key: str, value: str) -> None:
        with self._connect() as conn:
            conn.execute(
                """
                INSERT INTO gsm_setting (key, value) VALUES (?, ?)
                ON CONFLICT(key) DO UPDATE SET value = excluded.value
                """,
                (key, value),
            )
            conn.commit()
