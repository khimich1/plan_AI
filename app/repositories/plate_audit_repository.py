"""Тонкий репозиторий-обёртка для журнала переходов статусов плит.

Журнал ``plate_status_log`` создаётся в :func:`core.kp_db.init_schema`.
Репозиторий принимает уже открытый ``sqlite3.Cursor``, чтобы вставка
происходила в той же транзакции, что и собственно ``UPDATE kp_plates``
или ``INSERT completed_plates``. Это даёт атомарность: если основная
запись отменилась — отменится и audit-строка.
"""
from __future__ import annotations

from sqlite3 import Cursor


class PlateAuditRepository:
    """Запись событий изменения статуса плит в ``plate_status_log``."""

    @staticmethod
    def append(
        cur: Cursor,
        *,
        plate_id: int | None,
        kp_id: int,
        plate_name: str | None,
        plan_id: str | None,
        day_number: int | None,
        from_status: str | None,
        to_status: str,
        qty: int,
        reason: str,
        actor: str | None,
    ) -> None:
        """Добавляет запись о переходе статуса плиты.

        Принимает уже открытый ``cur`` — каллер отвечает за commit/rollback.
        """
        cur.execute(
            """
            INSERT INTO plate_status_log (
                plate_id, kp_id, plate_name, plan_id, day_number,
                from_status, to_status, qty, reason, actor
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                plate_id,
                int(kp_id),
                plate_name,
                plan_id,
                day_number,
                from_status,
                to_status,
                int(qty),
                reason,
                actor,
            ),
        )
