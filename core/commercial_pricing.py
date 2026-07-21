from __future__ import annotations

import logging
import sqlite3
from typing import Any

from core.cargo_delivery_pricing import delivery_service_charge_rub, total_order_cargo_weight_kg
from core.exceptions import PriceNotFoundError, UnpricedPlatesError
from core.price_db import length_m_to_price_length_dm

_log = logging.getLogger(__name__)

VAT_RATE = 0.22


def lookup_plate_price(
    length_m: float,
    width_m: float,
    load_class: int = 800,
    *,
    db_path: str,
) -> float:
    """Return plate unit price from ``pb.db`` or raise ``PriceNotFoundError``."""
    _ = width_m  # width is part of the public contract for callers / future pricing rules
    try:
        length_dm = length_m_to_price_length_dm(length_m)
        load_code = int(load_class) // 100
        con = sqlite3.connect(db_path)
        try:
            cur = con.cursor()
            row = cur.execute(
                "SELECT price FROM prices WHERE length_dm = ? AND load_code = ?",
                (length_dm, load_code),
            ).fetchone()
        finally:
            con.close()
        if row:
            return float(row[0])
        raise PriceNotFoundError(
            f"Цена не найдена: length_dm={length_dm}, load_code={load_code}"
        )
    except PriceNotFoundError:
        raise
    except Exception as exc:
        _log.warning(
            "Ошибка получения цены (length_m=%s, load_class=%s): %s",
            length_m,
            load_class,
            exc,
            exc_info=True,
        )
        raise PriceNotFoundError(
            f"Ошибка получения цены для length_m={length_m}, load_class={load_class}"
        ) from exc


def position_label(item: dict[str, Any]) -> str:
    name = str(item.get("name", "") or "").strip()
    if name:
        return name
    length_m = item.get("length_m", 0)
    width_m = item.get("width_m", 0)
    load_class = int(item.get("load_class", 800) or 800)
    return f"ПБ {length_m}-{width_m}-{load_class // 100}п"


def collect_unpriced_positions(
    order_data: list[dict[str, Any]],
    *,
    db_path: str,
) -> list[str]:
    unpriced: list[str] = []
    seen: set[str] = set()
    for item in order_data:
        if item.get("unit_price") is not None:
            continue
        length_m = float(item.get("length_m", 0) or 0)
        width_m = float(item.get("width_m", 0) or 0)
        load_class = int(item.get("load_class", 800) or 800)
        try:
            lookup_plate_price(length_m, width_m, load_class, db_path=db_path)
        except PriceNotFoundError:
            label = position_label(item)
            if label not in seen:
                seen.add(label)
                unpriced.append(label)
    return unpriced


def ensure_order_priced(
    order_data: list[dict[str, Any]],
    *,
    db_path: str,
) -> None:
    positions = collect_unpriced_positions(order_data, db_path=db_path)
    if positions:
        _log.warning("Непрорасценённые позиции: %s", positions)
        raise UnpricedPlatesError(positions)


def format_phone(phone: str) -> str:
    """Форматирует телефон в вид +7 (XXX) XXX-XX-XX."""
    if not phone:
        return ""

    digits = "".join(filter(str.isdigit, phone))
    if digits.startswith("7") and len(digits) == 11:
        return f"+7 ({digits[1:4]}) {digits[4:7]}-{digits[7:9]}-{digits[9:11]}"
    return digits


def calculate_total_cost(
    order_data: list[dict[str, Any]],
    discount_percent: float = 0,
    logistics_cost: float = 0,
    *,
    db_path: str,
) -> dict[str, Any]:
    """
    Рассчитывает общую стоимость заказа.

    unit_price в позициях считается уже с НДС. Скидка применяется к сумме плит.
    НДС для отображения: сумма плит после скидки * VAT_RATE.
    Итого к оплате: сумма плит после скидки + услуга по доставке грузов
    (стоимость рейса × ceil(масса заказа кг / 18600); в базу НДС по плитам не входит).

    subtotal = total_with_vat - vat_amount (согласованная разбивка для документов и архива).
    """
    ensure_order_priced(order_data, db_path=db_path)
    total_qty = 0
    plates_total_with_vat = 0.0
    dp = float(discount_percent or 0.0)
    dp = min(max(dp, 0.0), 100.0)
    discount_factor = 1.0 - dp / 100.0

    for item in order_data:
        qty = item.get("qty", 0)

        if "unit_price" in item and item["unit_price"] is not None:
            unit_price = item["unit_price"]
        else:
            length_m = item.get("length_m", 0)
            width_m = item.get("width_m", 0)
            load_class = item.get("load_class", 800)
            unit_price = lookup_plate_price(
                length_m, width_m, load_class, db_path=db_path
            )

        discounted_price = float(unit_price) * discount_factor
        item_cost = discounted_price * qty

        total_qty += qty
        plates_total_with_vat += item_cost

    trip_cost = max(0.0, float(logistics_cost or 0.0))
    cargo_kg = total_order_cargo_weight_kg(order_data)
    delivery_total = delivery_service_charge_rub(trip_cost, cargo_kg)
    vat_amount = round(plates_total_with_vat * VAT_RATE, 2)
    total_with_vat = round(plates_total_with_vat + delivery_total, 2)
    subtotal = round(total_with_vat - vat_amount, 2)

    return {
        "total_qty": total_qty,
        "subtotal": subtotal,
        "vat_amount": vat_amount,
        "total_with_vat": total_with_vat,
    }
