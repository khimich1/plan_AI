"""SQL persistence for the 1C counterparties directory."""

from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from typing import Any, Iterable

from core.kp_db_common import _connect
from core.kp_db_schema import ensure_schema


@dataclass(frozen=True)
class UpsertStats:
    added: int = 0
    updated: int = 0
    unchanged: int = 0


def _blank_to_none(value: Any) -> str | None:
    if value is None:
        return None
    stripped = str(value).strip()
    return stripped or None


def _as_flag(value: Any, *, default: bool = True) -> int:
    if value is None:
        return 1 if default else 0
    if isinstance(value, str):
        return 0 if value.strip().casefold() in {"0", "false", "нет", "no"} else 1
    return 1 if value else 0


def _normalize_record(raw: dict[str, Any], *, default_source: str) -> dict[str, Any]:
    name = (raw.get("name") or "").strip()
    code_1c = (raw.get("code_1c") or "").strip()
    return {
        "code_1c": code_1c,
        "name": name,
        "name_normalized": name.casefold(),
        "inn": _blank_to_none(raw.get("inn")),
        "kpp": _blank_to_none(raw.get("kpp")),
        "is_client": _as_flag(raw.get("is_client"), default=True),
        "is_active": _as_flag(raw.get("is_active"), default=True),
        "source": (raw.get("source") or default_source).strip(),
        "guid_1c": _blank_to_none(raw.get("guid_1c")),
    }


def _row_to_dict(row: sqlite3.Row | None) -> dict[str, Any] | None:
    if row is None:
        return None
    data = dict(row)
    data.pop("rank", None)
    return data


class CounterpartiesRepository:
    """Единственная точка доступа к таблице ``counterparties``."""

    def __init__(self, *, db_path: str) -> None:
        self.db_path = db_path

    def _connect(self) -> sqlite3.Connection:
        ensure_schema(self.db_path)
        conn = _connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    def get_by_id(self, counterparty_id: int) -> dict[str, Any] | None:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM counterparties WHERE id = ?",
                (int(counterparty_id),),
            ).fetchone()
        return _row_to_dict(row)

    def get_by_code_1c(self, code_1c: str) -> dict[str, Any] | None:
        code = (code_1c or "").strip()
        if not code:
            return None
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM counterparties WHERE code_1c = ?",
                (code,),
            ).fetchone()
        return _row_to_dict(row)

    def find_by_inn(self, inn: str) -> list[dict[str, Any]]:
        value = _blank_to_none(inn)
        if value is None:
            return []
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT * FROM counterparties WHERE inn = ? ORDER BY id",
                (value,),
            ).fetchall()
        return [dict(row) for row in rows]

    def list_codes(self) -> set[str]:
        with self._connect() as conn:
            rows = conn.execute("SELECT code_1c FROM counterparties").fetchall()
        return {str(row["code_1c"]) for row in rows}

    def search(
        self,
        q: str,
        *,
        limit: int = 10,
        clients_only: bool = True,
    ) -> list[dict[str, Any]]:
        needle = (q or "").strip().casefold()
        if not needle:
            return []
        digits = needle if needle.isdigit() else None
        inn_like = f"{digits}%" if digits is not None else ""
        name_prefix = f"{needle}%"
        name_substr = f"%{needle}%"
        sql = """
            SELECT
                id, code_1c, guid_1c, name, name_normalized, inn, kpp,
                is_client, is_active, source, created_at, updated_at,
                CASE
                    WHEN name_normalized LIKE ? THEN 1
                    WHEN name_normalized LIKE ? THEN 2
                    WHEN ? IS NOT NULL AND inn LIKE ? THEN 3
                    ELSE 4
                END AS rank
            FROM counterparties
            WHERE is_active = 1
              AND (? = 0 OR is_client = 1)
              AND (
                    name_normalized LIKE ?
                    OR (? IS NOT NULL AND inn LIKE ?)
              )
            ORDER BY rank, name_normalized
            LIMIT ?
        """
        params = (
            name_prefix,
            name_substr,
            digits,
            inn_like,
            1 if clients_only else 0,
            name_substr,
            digits,
            inn_like,
            int(limit),
        )
        with self._connect() as conn:
            rows = conn.execute(sql, params).fetchall()
        return [item for item in (_row_to_dict(row) for row in rows) if item is not None]

    def insert(
        self,
        *,
        code_1c: str,
        name: str,
        inn: str | None = None,
        kpp: str | None = None,
        is_client: bool = True,
        is_active: bool = True,
        source: str = "manual",
        guid_1c: str | None = None,
    ) -> dict[str, Any]:
        record = _normalize_record(
            {
                "code_1c": code_1c,
                "name": name,
                "inn": inn,
                "kpp": kpp,
                "is_client": is_client,
                "is_active": is_active,
                "source": source,
                "guid_1c": guid_1c,
            },
            default_source=source,
        )
        with self._connect() as conn:
            cur = conn.execute(
                """
                INSERT INTO counterparties (
                    code_1c, guid_1c, name, name_normalized, inn, kpp,
                    is_client, is_active, source
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    record["code_1c"],
                    record["guid_1c"],
                    record["name"],
                    record["name_normalized"],
                    record["inn"],
                    record["kpp"],
                    record["is_client"],
                    record["is_active"],
                    record["source"],
                ),
            )
            new_id = int(cur.lastrowid)
            conn.commit()
        row = self.get_by_id(new_id)
        assert row is not None
        return row

    def upsert(self, records: Iterable[dict[str, Any]]) -> UpsertStats:
        added = updated = unchanged = 0
        with self._connect() as conn:
            for raw in records:
                rec = _normalize_record(raw, default_source="import")
                if not rec["code_1c"] or not rec["name"]:
                    continue
                existing = conn.execute(
                    "SELECT * FROM counterparties WHERE code_1c = ?",
                    (rec["code_1c"],),
                ).fetchone()
                if existing is None:
                    conn.execute(
                        """
                        INSERT INTO counterparties (
                            code_1c, guid_1c, name, name_normalized, inn, kpp,
                            is_client, is_active, source
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            rec["code_1c"],
                            rec["guid_1c"],
                            rec["name"],
                            rec["name_normalized"],
                            rec["inn"],
                            rec["kpp"],
                            rec["is_client"],
                            rec["is_active"],
                            rec["source"],
                        ),
                    )
                    added += 1
                    continue
                same = (
                    existing["name"] == rec["name"]
                    and existing["inn"] == rec["inn"]
                    and existing["kpp"] == rec["kpp"]
                    and int(existing["is_client"]) == rec["is_client"]
                )
                if same:
                    unchanged += 1
                    continue
                conn.execute(
                    """
                    UPDATE counterparties
                    SET name = ?, name_normalized = ?, inn = ?, kpp = ?,
                        is_client = ?, updated_at = datetime('now')
                    WHERE id = ?
                    """,
                    (
                        rec["name"],
                        rec["name_normalized"],
                        rec["inn"],
                        rec["kpp"],
                        rec["is_client"],
                        int(existing["id"]),
                    ),
                )
                updated += 1
            conn.commit()
        return UpsertStats(added=added, updated=updated, unchanged=unchanged)
