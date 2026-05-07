from __future__ import annotations

import asyncio
import logging
import os
import re
from datetime import datetime, timedelta
from pathlib import Path
from typing import Literal

from app.core.settings import get_settings
from app.repositories.kp_archive_repository import ArchiveSection, KpArchiveRepository
from app.schemas.archive import (
    ArchiveFileKind,
    ArchiveOfferDetails,
    ArchiveOfferFinance,
    ArchiveOfferListItem,
    ArchivePlateItem,
)
from core.commercial_offer import generate_commercial_offer_pdf
from core.commercial_offer_xlsx import generate_commercial_offer_xlsx
from core.gantt_excel import create_gantt_excel


logger = logging.getLogger(__name__)

_MAX_TRACK_LENGTH_M = 101.0
_DAYS_PER_TRACK_FACTOR = 5.0


class ArchiveError(Exception):
    """Базовое исключение сервиса архива."""


class ArchiveNotFoundError(ArchiveError):
    """КП не найдено в БД."""


class ArchiveValidationError(ArchiveError):
    """Ошибка валидации бизнес-правил (например, недопустимый статус)."""


class ArchiveService:
    """Бизнес-логика раздела «Архив» (аналог bot/handlers/archive.py)."""

    def __init__(
        self,
        repository: KpArchiveRepository | None = None,
        outputs_dir: Path | None = None,
    ) -> None:
        settings = get_settings()
        self.repository = repository or KpArchiveRepository()
        self.outputs_dir = outputs_dir or settings.outputs_dir
        self.outputs_dir.mkdir(parents=True, exist_ok=True)

    # ---------- Списки и карточка ----------

    def list_offers(self, section: ArchiveSection) -> list[ArchiveOfferListItem]:
        raw_items = self.repository.list_by_section(section)
        items: list[ArchiveOfferListItem] = []
        for raw in raw_items:
            kp_id = int(raw.get("kp_id") or 0)
            completion = None
            if section in ("in_production", "completed"):
                try:
                    completion = float(
                        self.repository.get_completion_percentage(kp_id).get("percentage", 0.0)
                    )
                except Exception:
                    logger.exception("Ошибка получения %% выполнения для КП %s", kp_id)
                    completion = None
            items.append(
                ArchiveOfferListItem(
                    kp_id=kp_id,
                    creation_date=raw.get("creation_date"),
                    customer_name=raw.get("customer_name"),
                    manager_name=raw.get("manager_name"),
                    discount_percent=float(raw.get("discount_percent") or 0),
                    subtotal=float(raw.get("subtotal") or 0),
                    vat_amount=float(raw.get("vat_amount") or 0),
                    total_amount=float(raw.get("total_amount") or 0),
                    execution_terms=raw.get("execution_terms") or None,
                    status=raw.get("status") or None,
                    completion_percentage=completion,
                )
            )
        return items

    def get_details(self, kp_id: int) -> ArchiveOfferDetails:
        raw = self.repository.get_by_id(kp_id)
        if not raw:
            raise ArchiveNotFoundError(f"КП №{kp_id} не найдено")
        return self._to_details(raw)

    def search_by_number(self, kp_id: int) -> ArchiveOfferDetails | None:
        raw = self.repository.get_by_id(kp_id)
        if not raw:
            return None
        return self._to_details(raw)

    # ---------- Мутации ----------

    def update_discount(self, kp_id: int, discount: float) -> ArchiveOfferDetails:
        if not 0 <= discount <= 100:
            raise ArchiveValidationError("Процент скидки должен быть от 0 до 100")
        if not self.repository.update_discount(kp_id, discount):
            raise ArchiveNotFoundError(
                f"Не удалось обновить скидку. КП №{kp_id} не найдено или пустое."
            )
        return self.get_details(kp_id)

    def delete_offer(self, kp_id: int) -> None:
        if not self.repository.delete(kp_id):
            raise ArchiveNotFoundError(f"КП №{kp_id} уже удалено или не существует")

    def move_to_production(self, kp_id: int, terms_input: str) -> ArchiveOfferDetails:
        raw = self.repository.get_by_id(kp_id)
        if not raw:
            raise ArchiveNotFoundError(f"КП №{kp_id} не найдено")
        if raw.get("status") != "в архиве":
            raise ArchiveValidationError(
                "Перевести в производство можно только КП из раздела «в архиве»"
            )

        execution_terms = self._parse_execution_terms(terms_input)
        if not self.repository.update_execution_date(kp_id, execution_terms):
            raise ArchiveError(f"Не удалось сохранить срок для КП №{kp_id}")
        if not self.repository.update_status(kp_id, "в работе"):
            raise ArchiveError(f"Не удалось изменить статус для КП №{kp_id}")

        return self.get_details(kp_id)

    def estimate_production(self, kp_id: int) -> dict:
        raw = self.repository.get_by_id(kp_id)
        if not raw:
            raise ArchiveNotFoundError(f"КП №{kp_id} не найдено")
        plates = raw.get("plates") or []
        total_length = sum((p.get("length_m") or 0) * (p.get("qty") or 1) for p in plates)
        estimated_tracks = max(1, int(round(total_length / _MAX_TRACK_LENGTH_M + 0.5)))
        estimated_days = max(1, int(round(estimated_tracks / _DAYS_PER_TRACK_FACTOR + 0.5)))
        return {
            "total_length_m": total_length,
            "estimated_tracks": estimated_tracks,
            "estimated_days": estimated_days,
        }

    # ---------- Документы ----------

    async def generate_document(self, kp_id: int, kind: ArchiveFileKind) -> Path:
        raw = self.repository.get_by_id(kp_id)
        if not raw:
            raise ArchiveNotFoundError(f"КП №{kp_id} не найдено")

        order_data = self._order_data_from_kp_info(raw)
        if not order_data:
            raise ArchiveValidationError("В КП нет позиций для формирования документа")

        offer_number = str(kp_id)
        offer_date = raw.get("creation_date") or datetime.now().strftime("%d.%m.%Y")
        customer_name = raw.get("customer_name")
        manager_name = raw.get("manager_name")
        discount_percent = float(raw.get("discount_percent") or 0)

        if kind == "pdf":
            buffer = await asyncio.to_thread(
                generate_commercial_offer_pdf,
                order_data,
                offer_number,
                offer_date,
                customer_name=customer_name,
                manager_name=manager_name,
                manager_phone=None,
                manager_email=None,
                discount_percent=discount_percent,
                kp_db_id=kp_id,
            )
            filename = f"КП_{kp_id}.pdf"
        elif kind == "xlsx":
            buffer = await asyncio.to_thread(
                generate_commercial_offer_xlsx,
                order_data,
                offer_number,
                offer_date,
                customer_name=customer_name,
                manager_name=manager_name,
                manager_phone=None,
                manager_email=None,
                discount_percent=discount_percent,
                delivery_conditions=raw.get("delivery_conditions"),
                payment_conditions=raw.get("payment_conditions"),
                kp_db_id=kp_id,
            )
            filename = f"КП_{kp_id}.xlsx"
        else:
            raise ArchiveValidationError(f"Неподдерживаемый тип файла: {kind}")

        target_path = self.outputs_dir / filename
        await asyncio.to_thread(_write_bytes, target_path, buffer.getvalue())
        return target_path

    # ---------- «Актуальный план» (Gantt) ----------

    async def build_current_plan_gantt(self) -> Path:
        """
        Собирает сводную диаграмму Ганта по всем сохранённым планам.
        Импорт plan_manager отложен, чтобы бэкенд запускался без bot-окружения.
        """
        from app.planning.plan_manager import get_all_plans_gantt_data  # локальный импорт

        gantt_data = await asyncio.to_thread(get_all_plans_gantt_data)
        if not gantt_data:
            raise ArchiveValidationError(
                "Нет сохранённых планов для создания диаграммы."
            )

        gantt_path = await asyncio.to_thread(
            create_gantt_excel,
            all_tracks_list=gantt_data["all_tracks"],
            tracks_count=3,
            plate_lookup_exact=gantt_data["plate_lookup_exact"],
            plate_lookup_by_length=gantt_data["plate_lookup_by_length"],
            output_dir=str(self.outputs_dir),
            start_date=gantt_data["earliest_start_date"],
        )
        if not gantt_path or not os.path.exists(gantt_path):
            raise ArchiveError("Не удалось создать диаграмму Ганта")
        return Path(gantt_path)

    # ---------- helpers ----------

    def _to_details(self, raw: dict) -> ArchiveOfferDetails:
        plates = [self._plate_item(p) for p in (raw.get("plates") or [])]
        kp_id = int(raw.get("kp_id") or 0)
        completion = None
        if raw.get("status") in ("в работе", "выполнено"):
            try:
                completion = float(
                    self.repository.get_completion_percentage(kp_id).get("percentage", 0.0)
                )
            except Exception:
                logger.exception("Ошибка получения %% выполнения для КП %s", kp_id)

        return ArchiveOfferDetails(
            kp_id=kp_id,
            creation_date=raw.get("creation_date"),
            customer_name=raw.get("customer_name"),
            manager_name=raw.get("manager_name"),
            status=raw.get("status"),
            execution_terms=raw.get("execution_terms") or None,
            delivery_conditions=raw.get("delivery_conditions") or None,
            payment_conditions=raw.get("payment_conditions") or None,
            finance=ArchiveOfferFinance(
                subtotal=float(raw.get("subtotal") or 0),
                vat_amount=float(raw.get("vat_amount") or 0),
                total_amount=float(raw.get("total_amount") or 0),
                discount_percent=float(raw.get("discount_percent") or 0),
            ),
            plates=plates,
            completion_percentage=completion,
        )

    @staticmethod
    def _plate_item(raw: dict) -> ArchivePlateItem:
        return ArchivePlateItem(
            position_number=raw.get("position_number"),
            plate_name=raw.get("plate_name") or "",
            length_m=_nullable_float(raw.get("length_m")),
            width_m=_nullable_float(raw.get("width_m")),
            load_class=_nullable_int(raw.get("load_class")),
            qty=int(raw.get("qty") or 0),
            unit_price=_nullable_float(raw.get("unit_price")),
            discounted_price=_nullable_float(raw.get("discounted_price")),
            unit_weight=_nullable_float(raw.get("unit_weight")),
            total_weight=_nullable_float(raw.get("total_weight")),
            status=raw.get("status") or None,
        )

    @staticmethod
    def _order_data_from_kp_info(kp_info: dict) -> list[dict]:
        """Повторяет логику bot/handlers/archive._order_data_from_kp_info."""
        plates = kp_info.get("plates") or []
        discount = kp_info.get("discount_percent") or 0
        factor = 1.0 - (discount / 100.0)
        if factor <= 0:
            factor = 1.0
        result = []
        for p in plates:
            unit_price = p.get("unit_price")
            if unit_price is None or (isinstance(unit_price, (int, float)) and unit_price <= 0):
                discounted_price = p.get("discounted_price") or 0
                unit_price = discounted_price / factor
            qty = p.get("qty") or 0
            total_weight = p.get("total_weight")
            unit_weight = p.get("unit_weight")
            weight = (
                total_weight
                if total_weight is not None and total_weight > 0
                else (unit_weight or 0) * qty
            )
            result.append(
                {
                    "name": p.get("plate_name") or "",
                    "length_m": p.get("length_m") or 0,
                    "width_m": p.get("width_m") or 0,
                    "qty": qty,
                    "load_class": p.get("load_class") or 800,
                    "unit_price": float(unit_price),
                    "weight": weight or 0,
                }
            )
        return result

    @staticmethod
    def _parse_execution_terms(raw: str) -> str:
        value = (raw or "").strip()
        if not value:
            raise ArchiveValidationError("Укажите срок изготовления")

        for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
            try:
                dt = datetime.strptime(value, fmt)
                return dt.strftime("%d.%m.%Y")
            except ValueError:
                continue

        match_days = re.search(r"(\d+)\s*(?:дн|день|дней|day|days)", value, re.IGNORECASE)
        if match_days:
            days = int(match_days.group(1))
            return (datetime.now() + timedelta(days=days)).strftime("%d.%m.%Y")

        match_weeks = re.search(r"(\d+)\s*(?:нед|недел|недели|week|weeks)", value, re.IGNORECASE)
        if match_weeks:
            weeks = int(match_weeks.group(1))
            return (datetime.now() + timedelta(weeks=weeks)).strftime("%d.%m.%Y")

        raise ArchiveValidationError(
            "Не удалось распознать срок. Используйте формат ДД.ММ.ГГГГ, ГГГГ-ММ-ДД, 'N дней' или 'N недель'."
        )


def _nullable_float(value: object) -> float | None:
    if value is None:
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _nullable_int(value: object) -> int | None:
    if value is None:
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _write_bytes(path: Path, data: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)
