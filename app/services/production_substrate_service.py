"""Service: substrate recommendations from a full-backlog optimizer pass."""

from __future__ import annotations

from collections import defaultdict
from datetime import date, datetime
from typing import Any, Mapping, Sequence

from app.core.settings import get_settings
from app.repositories.kp_repository import KpRepository
from core.domain.plate_order import normalize_load_code
from core.optimization import optimize_with_cascading_longitudinal_cuts
from core.optimization.result_contract import (
    OPT_ERROR_MESSAGE_KEY,
    OPT_STATUS_KEY,
)
from core.production.substrate import (
    SubstrateRecommendation,
    extract_substrate_recommendations,
)
from core.production.urgent import collect_urgent_positions

__all__ = [
    "ProductionSubstrateError",
    "ProductionSubstrateService",
    "SubstrateRecommendation",
]


class ProductionSubstrateError(RuntimeError):
    """Domain error when substrate analysis cannot produce recommendations."""


class ProductionSubstrateService:
    """Runs a synchronous backlog optimize and extracts cross-KP substrates."""

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

    def find_substrate_recommendations(
        self,
        *,
        urgent_plate_ids: Sequence[int],
        deadline_until: date,
        first_fill_target_date: date,
        now: datetime | None = None,
    ) -> list[SubstrateRecommendation]:
        """Recommend late plates that the optimizer nests under urgent primaries.

        ``deadline_until`` is accepted for API parity (caller already filtered
        urgent ids); substrate ``needed_by`` is computed for the full backlog
        without that filter (deadlines may fall after ``deadline_until``).
        """
        _ = deadline_until  # API parity; deadlines built with date.max below

        backlog = self._load_backlog_plates()
        if not backlog:
            return []

        orders_2d = self._build_orders_2d(backlog)
        if not orders_2d:
            return []

        result = optimize_with_cascading_longitudinal_cuts(orders_2d=orders_2d)
        if not isinstance(result, dict):
            raise ProductionSubstrateError(
                "Оптимизатор не вернул результат для анализа подложек."
            )
        if result.get(OPT_STATUS_KEY) == "error":
            message = result.get(OPT_ERROR_MESSAGE_KEY) or "ошибка оптимизатора"
            raise ProductionSubstrateError(
                f"Оптимизатор вернул ошибку при анализе подложек: {message}"
            )

        plate_id_by_kp_name = {
            (int(p["kp_id"]), str(p["plate_name"])): int(p["plate_id"])
            for p in backlog
        }
        deadline_by_plate_id = self._deadline_by_plate_id(backlog, now=now)

        return extract_substrate_recommendations(
            result,
            plate_id_by_kp_name=plate_id_by_kp_name,
            urgent_plate_ids=urgent_plate_ids,
            deadline_by_plate_id=deadline_by_plate_id,
            first_fill_target_date=first_fill_target_date,
        )

    def _load_backlog_plates(self) -> list[dict[str, Any]]:
        plates: list[dict[str, Any]] = []
        for kp in self.kp_repository.list_kps_in_production():
            kp_id = int(kp["kp_id"])
            for plate in kp.get("plates") or []:
                plate_id = int(plate["id"])
                qty_remaining = int(
                    self.kp_repository.get_plate_qty_remaining(plate_id)
                )
                if qty_remaining <= 0:
                    continue
                plates.append(
                    {
                        "plate_id": plate_id,
                        "kp_id": kp_id,
                        "plate_name": plate.get("plate_name") or "",
                        "length_m": float(plate.get("length_m") or 0.0),
                        "width_m": float(plate.get("width_m") or 0.0),
                        "load_class": plate.get("load_class"),
                        "qty_remaining": qty_remaining,
                        "execution_terms": str(kp.get("execution_terms") or ""),
                    }
                )
        return plates

    @staticmethod
    def _build_orders_2d(
        backlog: Sequence[Mapping[str, Any]],
    ) -> list[dict[str, Any]]:
        orders: list[dict[str, Any]] = []
        for plate in backlog:
            length_m = float(plate["length_m"])
            width_m = float(plate["width_m"])
            qty = int(plate["qty_remaining"])
            if length_m <= 0 or qty <= 0:
                continue
            width_mm = int(round(width_m * 1000))
            if width_mm < 1:
                continue
            orders.append(
                {
                    "length": length_m,
                    "width": width_mm,
                    "qty": qty,
                    "load_code": normalize_load_code(
                        plate.get("load_class"), default=8
                    ),
                    "kp_id": int(plate["kp_id"]),
                    "plate_name": str(plate.get("plate_name") or ""),
                    "reinforcement": 0,
                    "concrete_grade": None,
                }
            )
        return orders

    def _deadline_by_plate_id(
        self,
        backlog: Sequence[Mapping[str, Any]],
        *,
        now: datetime | None = None,
    ) -> dict[int, date]:
        """Deadlines for all backlog plates (no ``deadline_until`` filter)."""
        plates_for_urgent: list[dict[str, Any]] = []
        kp_meta: dict[int, dict[str, str]] = {}
        for plate in backlog:
            kp_id = int(plate["kp_id"])
            kp_meta.setdefault(
                kp_id,
                {"execution_terms": str(plate.get("execution_terms") or "")},
            )
            plates_for_urgent.append(
                {
                    "plate_id": int(plate["plate_id"]),
                    "kp_id": kp_id,
                    "plate_name": str(plate.get("plate_name") or ""),
                    "qty_remaining": int(plate.get("qty_remaining") or 0),
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

        positions = collect_urgent_positions(
            plates_for_urgent,
            batches_by_plate,
            kp_meta,
            date.max,
            now=now,
        )
        return {int(p.plate_id): p.deadline for p in positions}
