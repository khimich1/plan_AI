"""Сверка очереди guid_demand (pb.db) с КП (plita.db) после правки состава КП.

Вызывается строго после того, как КП записан в plita.db. Наружу не бросает
никогда: обрыв оставляет лишнюю строку очереди, следующая сверка её исправит.
"""

from __future__ import annotations

import logging
import sqlite3
from pathlib import Path

from core.guid_demand_reconcile import ReconcileReport, reconcile_guid_demand

logger = logging.getLogger(__name__)

# Залипший писатель pb.db не должен вешать запрос сохранения/просмотра.
_PB_TIMEOUT_SEC = 3.0


def sync_guid_queue_with_kp(
    pb_db_path: str | Path,
    plita_db_path: str | Path,
) -> ReconcileReport | None:
    """Привести очередь к составу КП. ``None`` — сверка не выполнена (ошибка/нет базы).

    plita.db открывается только на чтение; pb.db коммитится, только если сверка
    что-то изменила.
    """
    try:
        pb_path = Path(pb_db_path)
        plita_path = Path(plita_db_path)
        if not pb_path.is_file() or not plita_path.is_file():
            logger.debug("Сверка очереди GUID пропущена: нет pb.db или plita.db")
            return None
        pb_conn = sqlite3.connect(str(pb_path), timeout=_PB_TIMEOUT_SEC)
        try:
            plita_conn = sqlite3.connect(
                f"{plita_path.resolve().as_uri()}?mode=ro", uri=True
            )
            try:
                report = reconcile_guid_demand(pb_conn, plita_conn)
            finally:
                plita_conn.close()
            if report.has_changes:
                pb_conn.commit()
            else:
                pb_conn.rollback()
            return report
        finally:
            pb_conn.close()
    except Exception:
        logger.exception("Не удалось сверить очередь «Завести в 1С» с КП")
        return None


__all__ = ["sync_guid_queue_with_kp"]
