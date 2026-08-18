"""Service layer for urgent production positions (delivery batches + KP terms)."""

from __future__ import annotations

from collections import defaultdict
from datetime import date, datetime

from app.core.settings import get_settings
from app.repositories.kp_repository import KpRepository
from core.production.urgent import UrgentPosition, collect_urgent_positions

__all__ = ["ProductionUrgentService", "UrgentPosition"]


class ProductionUrgentService:
    """Assembles DB backlog and delegates aggregation to ``collect_urgent_positions``."""

    def __init__(
        self,
        *,
        db_path: str | None = None,
        kp_repository: KpRepository | None = None,
    ) -> None:
        if kp_repository is not None:
            self.kp_repository = kp_repository
            self.db_path = kp_repository.db_path
        else:
            self.db_path = str(db_path or get_settings().plita_db_path)
            self.kp_repository = KpRepository(db_path=self.db_path)

    def list_urgent_positions(
        self,
        *,
        deadline_until: date,
        now: datetime | None = None,
    ) -> list[UrgentPosition]:
        kps = self.kp_repository.list_kps_in_production()

        plates: list[dict] = []
        kp_meta: dict[int, dict[str, str]] = {}
        for kp in kps:
            kp_id = int(kp["kp_id"])
            kp_meta[kp_id] = {"execution_terms": str(kp.get("execution_terms") or "")}
            for plate in kp.get("plates") or []:
                plate_id = int(plate["id"])
                plates.append(
                    {
                        "plate_id": plate_id,
                        "kp_id": kp_id,
                        "plate_name": plate.get("plate_name") or "",
                        "qty_remaining": self.kp_repository.get_plate_qty_remaining(
                            plate_id
                        ),
                    }
                )

        batches_by_plate: dict[int, list[dict]] = defaultdict(list)
        for row in self.kp_repository.list_delivery_batch_items_for_in_production_plates():
            batches_by_plate[int(row["plate_id"])].append(
                {
                    "produce_by": row["produce_by"],
                    "qty": row["qty"],
                    "batch_name": row["batch_name"],
                }
            )

        return collect_urgent_positions(
            plates,
            batches_by_plate,
            kp_meta,
            deadline_until,
            now=now,
        )
