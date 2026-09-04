"""Сброс ГСМ к imported-якорям (dev / test-run)."""

from __future__ import annotations

from pathlib import Path

from app.schemas.gsm import GsmAnchorOut, GsmResetToAnchorsReport
from core.gsm.reset_to_anchors import ResetGsmError, run_reset


class GsmResetError(Exception):
    def __init__(self, message: str, *, code: str = "gsm_reset_error") -> None:
        super().__init__(message)
        self.code = code


class GsmResetService:
    def __init__(self, *, db_path: str) -> None:
        self.db_path = db_path

    def reset_to_anchors(self) -> GsmResetToAnchorsReport:
        try:
            result = run_reset(db_path=Path(self.db_path), apply=True)
        except ResetGsmError as exc:
            message = str(exc)
            code = (
                "gsm_reset_no_anchors"
                if "нет imported-якоря" in message
                else "gsm_reset_error"
            )
            raise GsmResetError(message, code=code) from exc

        plan = result.plan
        if result.backup_path is None:
            raise GsmResetError("backup was not created", code="gsm_reset_error")

        return GsmResetToAnchorsReport(
            backup_path=str(result.backup_path),
            anchors_kept=len(plan.anchors),
            waybills_deleted=plan.waybills_to_delete,
            transactions_deleted=plan.txs_total,
            import_batches_deleted=plan.batches_total,
            anchors=[
                GsmAnchorOut(
                    waybill_id=anchor.waybill_id,
                    vehicle_id=anchor.vehicle_id,
                    vehicle_name=anchor.name,
                    plate_number=anchor.plate_number,
                    date=anchor.date,
                    odometer_end=anchor.odometer_end,
                    fuel_end=anchor.fuel_end,
                )
                for anchor in plan.anchors
            ],
        )
