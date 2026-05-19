from __future__ import annotations

from datetime import datetime, timedelta
from typing import Any

from app.repositories.kp_repository import KpRepository
from app.schemas.offers import CreateOfferRequest
from core import kp_db
from core.execution_terms import parse_execution_terms_to_datetime
from core.commercial_offer import generate_commercial_offer_pdf
from core.commercial_offer_xlsx import generate_commercial_offer_xlsx


class OffersService:
    def __init__(self) -> None:
        self.kp_repository = KpRepository()

    def list_offers(
        self,
        *,
        status: str = "all",
        limit: int = 200,
        kp_id: int | None = None,
    ) -> list[dict[str, Any]]:
        if kp_id is not None:
            item = self.kp_repository.get_offer(kp_id)
            if not item:
                return []
            return [self._to_offer_summary(item)]

        grouped = self.kp_repository.list_offers_grouped()
        groups_map = {
            "archived": grouped.get("archived", []),
            "in_production": grouped.get("in_production", []),
            "completed": grouped.get("completed", []),
        }
        if status == "all":
            items = groups_map["archived"] + groups_map["in_production"] + groups_map["completed"]
        else:
            items = groups_map.get(status, [])
        items = sorted(items, key=lambda x: int(x.get("kp_id") or 0), reverse=True)
        return [self._to_offer_summary(item) for item in items[: max(limit, 1)]]

    def get_offer(self, kp_id: int) -> dict[str, Any] | None:
        item = self.kp_repository.get_offer(kp_id)
        if not item:
            return None
        return self._to_offer_details(item)

    def create_offer(self, payload: CreateOfferRequest) -> dict[str, Any]:
        execution_terms = None
        status = "в архиве"
        used_default_execution_terms = False
        if payload.save_mode == "work":
            execution_terms, used_default_execution_terms = self._parse_execution_terms(payload.execution_terms_input)
            status = "в работе"

        kp_id = kp_db.save_kp_to_db(
            creation_date=payload.creation_date,
            order_data=[item.model_dump() for item in payload.order_data],
            customer_name=payload.customer_name,
            manager_name=payload.manager_name,
            discount_percent=payload.discount_percent,
            delivery_conditions=payload.delivery_conditions,
            payment_conditions=payload.payment_conditions,
            execution_terms=execution_terms,
            status=status,
            db_path=self.kp_repository.db_path,
        )
        created = self.kp_repository.get_offer(kp_id)
        if not created:
            return {"kp_id": kp_id, "status": status, "execution_terms": execution_terms}
        return {
            "kp_id": kp_id,
            "status": status,
            "execution_terms": execution_terms,
            "used_default_execution_terms": used_default_execution_terms,
            "offer": self._to_offer_summary(created),
        }

    def update_discount(self, kp_id: int, discount_percent: float) -> dict[str, Any] | None:
        if not self.kp_repository.update_offer_discount(kp_id, discount_percent):
            return None
        item = self.kp_repository.get_offer(kp_id)
        if not item:
            return None
        return self._to_offer_details(item)

    def move_to_production(self, kp_id: int, execution_terms_input: str) -> dict[str, Any]:
        item = self.kp_repository.get_offer(kp_id)
        if not item:
            raise ValueError("not_found")
        if item.get("status") != "в архиве":
            raise ValueError("invalid_status")
        execution_terms, used_default = self._parse_execution_terms(execution_terms_input)
        if not self.kp_repository.update_offer_execution_date(kp_id, execution_terms):
            raise ValueError("update_execution_date_failed")
        if not self.kp_repository.update_offer_status(kp_id, "в работе"):
            raise ValueError("update_status_failed")
        updated = self.kp_repository.get_offer(kp_id)
        if not updated:
            raise ValueError("not_found")
        return {
            "kp_id": kp_id,
            "execution_terms": execution_terms,
            "used_default_execution_terms": used_default,
            "offer": self._to_offer_summary(updated),
        }

    def delete_offer(self, kp_id: int) -> bool:
        return self.kp_repository.delete_offer(kp_id)

    def generate_pdf(self, kp_id: int) -> tuple[str, bytes]:
        offer = self.kp_repository.get_offer(kp_id)
        if not offer:
            raise ValueError("not_found")
        order_data = self._order_data_from_offer(offer)
        if not order_data:
            raise ValueError("empty_order_data")
        offer_date = offer.get("creation_date") or datetime.now().strftime("%d.%m.%Y")
        pdf_buffer = generate_commercial_offer_pdf(
            order_data=order_data,
            offer_number=str(kp_id),
            offer_date=offer_date,
            customer_name=offer.get("customer_name"),
            manager_name=offer.get("manager_name"),
            manager_phone=None,
            manager_email=None,
            discount_percent=float(offer.get("discount_percent") or 0),
            kp_db_id=kp_id,
            delivery_conditions=offer.get("delivery_conditions"),
            payment_conditions=offer.get("payment_conditions"),
        )
        return (f"KP_{kp_id}.pdf", pdf_buffer.getvalue())

    def generate_xlsx(self, kp_id: int) -> tuple[str, bytes]:
        offer = self.kp_repository.get_offer(kp_id)
        if not offer:
            raise ValueError("not_found")
        order_data = self._order_data_from_offer(offer)
        if not order_data:
            raise ValueError("empty_order_data")
        offer_date = offer.get("creation_date") or datetime.now().strftime("%d.%m.%Y")
        xlsx_buffer = generate_commercial_offer_xlsx(
            order_data=order_data,
            offer_number=str(kp_id),
            offer_date=offer_date,
            customer_name=offer.get("customer_name"),
            manager_name=offer.get("manager_name"),
            manager_phone=None,
            manager_email=None,
            discount_percent=float(offer.get("discount_percent") or 0),
            delivery_conditions=offer.get("delivery_conditions"),
            payment_conditions=offer.get("payment_conditions"),
            kp_db_id=kp_id,
        )
        return (f"KP_{kp_id}.xlsx", xlsx_buffer.getvalue())

    def _parse_execution_terms(self, raw_terms: str) -> tuple[str, bool]:
        text = (raw_terms or "").strip()
        deadline_date = parse_execution_terms_to_datetime(text)
        used_default = False
        if not deadline_date:
            deadline_date = datetime.now() + timedelta(days=14)
            used_default = True
        return deadline_date.strftime("%d.%m.%Y"), used_default

    def _to_offer_summary(self, item: dict[str, Any]) -> dict[str, Any]:
        kp_id = int(item.get("kp_id") or 0)
        summary = {
            "kp_id": kp_id,
            "creation_date": item.get("creation_date"),
            "customer_name": item.get("customer_name"),
            "manager_name": item.get("manager_name"),
            "discount_percent": float(item.get("discount_percent") or 0),
            "subtotal": float(item.get("subtotal") or 0),
            "vat_amount": float(item.get("vat_amount") or 0),
            "total_amount": float(item.get("total_amount") or 0),
            "delivery_conditions": item.get("delivery_conditions"),
            "payment_conditions": item.get("payment_conditions"),
            "execution_terms": item.get("execution_terms"),
            "status": item.get("status") or "в работе",
        }
        if summary["status"] in {"в работе", "выполнено"}:
            completion = self.kp_repository.get_completion_percentage(kp_id)
            summary["completion_percentage"] = completion.get("percentage", 0)
        else:
            summary["completion_percentage"] = 0
        return summary

    def _to_offer_details(self, item: dict[str, Any]) -> dict[str, Any]:
        details = self._to_offer_summary(item)
        details["plates"] = item.get("plates", [])
        return details

    def _order_data_from_offer(self, offer: dict[str, Any]) -> list[dict[str, Any]]:
        order_data: list[dict[str, Any]] = []
        for plate in offer.get("plates", []):
            qty = plate.get("qty") or 0
            if qty <= 0:
                continue
            unit_price = plate.get("unit_price")
            if unit_price is None:
                unit_price = plate.get("discounted_price") or 0
            weight = plate.get("total_weight")
            if not weight:
                unit_weight = plate.get("unit_weight") or 0
                weight = unit_weight * qty
            order_data.append(
                {
                    "name": plate.get("plate_name") or "",
                    "length_m": plate.get("length_m") or 0,
                    "width_m": plate.get("width_m") or 0,
                    "qty": qty,
                    "load_class": plate.get("load_class") or 800,
                    "unit_price": float(unit_price or 0),
                    "weight": float(weight or 0),
                    "length_dm_raw": plate.get("length_dm_raw") or "",
                    "nomenclature_id": plate.get("nomenclature_id"),
                }
            )
        return order_data
