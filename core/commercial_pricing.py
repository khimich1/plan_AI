from __future__ import annotations

import logging
import sqlite3
from typing import Any, Mapping

from core.cargo_delivery_pricing import (
    cargo_delivery_trips_count,
    delivery_service_charge_rub,
    total_order_cargo_weight_kg,
)
from core.exceptions import PriceNotFoundError, UnpricedPlatesError
from core.bridge_pile_price_db import get_bridge_pile_price
from core.fbs_price_db import get_fbs_price
from core.march_price_db import get_march_price, normalize_march_mark
from core.pile_price_db import get_pile_price
from core.price_db import length_m_to_price_length_dm
from core.step_price_db import get_step_price, normalize_step_mark

_log = logging.getLogger(__name__)

VAT_RATE = 0.22


def _resolve_pile_catalog_db_path(explicit: str | None, fallback_db_path: str) -> str:
    """Каталог свай живёт в plita.db, не в pb.db (прайс)."""
    if explicit:
        return explicit
    try:
        from core.config.settings import get_settings

        return str(get_settings().plita_db_path)
    except Exception:
        return fallback_db_path


def _compute_pile_delivery_breakdown(
    order_data: list[dict[str, Any]],
    *,
    overrides: dict[str, int] | None,
    catalog_db_path: str,
):
    from core.pile_catalog import load_pile_catalog, resolve_catalog_for_mark
    from core.pile_trip_pricing import compute_pile_trips

    entries = load_pile_catalog(catalog_db_path)
    return compute_pile_trips(
        order_data,
        overrides,
        lambda mark: resolve_catalog_for_mark(mark, entries),
    )


def lookup_pile_price(
    mark: str,
    concrete_grade: str,
    *,
    db_path: str,
) -> float:
    """Return pile unit price from ``pb.db`` or raise ``PriceNotFoundError``."""
    price = get_pile_price(mark, concrete_grade, db_path)
    if price is None:
        raise PriceNotFoundError(
            f"Свая не найдена в прайсе: {mark}, {concrete_grade}"
        )
    return price


def lookup_step_price(
    mark: str,
    *,
    db_path: str,
) -> float:
    """Return stair-step unit price from ``pb.db`` or raise ``PriceNotFoundError``."""
    price = get_step_price(normalize_step_mark(mark), db_path)
    if price is None:
        raise PriceNotFoundError(f"Ступень не найдена в прайсе: {mark}")
    return price


def lookup_march_price(
    mark: str,
    concrete_grade: str,
    *,
    db_path: str,
) -> float:
    """Return stair-march unit price from ``pb.db`` or raise ``PriceNotFoundError``."""
    price = get_march_price(normalize_march_mark(mark), concrete_grade, db_path)
    if price is None:
        raise PriceNotFoundError(
            f"Марш не найден в прайсе: {mark}, {concrete_grade}"
        )
    return price


def lookup_bridge_pile_price(
    mark: str,
    concrete_grade: str,
    *,
    db_path: str,
) -> float:
    """Return bridge-pile unit price from ``pb.db`` or raise ``PriceNotFoundError``."""
    price = get_bridge_pile_price(mark, concrete_grade, db_path)
    if price is None:
        raise PriceNotFoundError(
            f"Мостовая свая не найдена в прайсе: {mark}, {concrete_grade}"
        )
    return price


def lookup_fbs_price(
    mark: str,
    concrete_grade: str,
    *,
    db_path: str,
) -> float:
    """Return FBS unit price from ``pb.db`` or raise ``PriceNotFoundError``."""
    price = get_fbs_price(mark, concrete_grade, db_path)
    if price is None:
        raise PriceNotFoundError(
            f"ФБС не найден в прайсе: {mark}, {concrete_grade}"
        )
    return price


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
        if row is None or float(row[0]) <= 0:
            raise PriceNotFoundError(
                f"Цена не найдена: length_dm={length_dm}, load_code={load_code}"
            )
        return float(row[0])
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


def _is_pile_item(item: dict[str, Any]) -> bool:
    return str(item.get("product_kind", "") or "").lower() == "pile"


def _is_step_item(item: dict[str, Any]) -> bool:
    return str(item.get("product_kind", "") or "").lower() == "step"


def _is_march_item(item: dict[str, Any]) -> bool:
    return str(item.get("product_kind", "") or "").lower() == "march"


def _is_bridge_pile_item(item: dict[str, Any]) -> bool:
    return str(item.get("product_kind", "") or "").lower() == "bridge_pile"


def _is_fbs_item(item: dict[str, Any]) -> bool:
    return str(item.get("product_kind", "") or "").lower() == "fbs"


def is_pile_order(order_data: list[dict[str, Any]]) -> bool:
    """True when every line is a pile position (empty list → False)."""
    return bool(order_data) and all(_is_pile_item(item) for item in order_data)


def is_step_order(order_data: list[dict[str, Any]]) -> bool:
    """True when every line is a stair-step position (empty list → False)."""
    return bool(order_data) and all(_is_step_item(item) for item in order_data)


def is_march_order(order_data: list[dict[str, Any]]) -> bool:
    """True when every line is a stair-march position (empty list → False)."""
    return bool(order_data) and all(_is_march_item(item) for item in order_data)


def is_bridge_pile_order(order_data: list[dict[str, Any]]) -> bool:
    """True when every line is a bridge-pile position (empty list → False)."""
    return bool(order_data) and all(_is_bridge_pile_item(item) for item in order_data)


def is_fbs_order(order_data: list[dict[str, Any]]) -> bool:
    """True when every line is an FBS position (empty list → False)."""
    return bool(order_data) and all(_is_fbs_item(item) for item in order_data)


def position_label(item: dict[str, Any]) -> str:
    if _is_pile_item(item):
        mark = str(item.get("mark") or item.get("name") or "").strip()
        grade = str(item.get("concrete_grade") or "B25").strip()
        if mark:
            return f"{mark} ({grade})"
        return f"Свая ({grade})"

    if _is_bridge_pile_item(item):
        mark = str(item.get("mark") or item.get("name") or "").strip()
        grade = str(item.get("concrete_grade") or "B25").strip()
        if mark:
            return f"{mark} ({grade})"
        return f"Мостовая свая ({grade})"

    if _is_fbs_item(item):
        mark = str(item.get("mark") or item.get("name") or "").strip()
        grade = str(item.get("concrete_grade") or "B25").strip()
        if mark:
            return f"{mark} ({grade})"
        return f"ФБС ({grade})"

    if _is_march_item(item):
        mark = str(item.get("mark") or item.get("name") or "").strip()
        grade = str(item.get("concrete_grade") or "B25").strip()
        if mark:
            return f"{mark} ({grade})"
        return f"Марш ({grade})"

    if _is_step_item(item):
        mark = str(item.get("mark") or item.get("name") or "").strip()
        return mark or "Ступень"

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
        try:
            if _is_pile_item(item):
                mark = str(item.get("mark") or item.get("name") or "").strip()
                grade = str(item.get("concrete_grade") or "B25").strip()
                lookup_pile_price(mark, grade, db_path=db_path)
            elif _is_bridge_pile_item(item):
                mark = str(item.get("mark") or item.get("name") or "").strip()
                grade = str(item.get("concrete_grade") or "B25").strip()
                lookup_bridge_pile_price(mark, grade, db_path=db_path)
            elif _is_fbs_item(item):
                mark = str(item.get("mark") or item.get("name") or "").strip()
                grade = str(item.get("concrete_grade") or "B25").strip()
                lookup_fbs_price(mark, grade, db_path=db_path)
            elif _is_march_item(item):
                mark = str(item.get("mark") or item.get("name") or "").strip()
                grade = str(item.get("concrete_grade") or "B25").strip()
                lookup_march_price(mark, grade, db_path=db_path)
            elif _is_step_item(item):
                mark = str(item.get("mark") or item.get("name") or "").strip()
                lookup_step_price(mark, db_path=db_path)
            else:
                length_m = float(item.get("length_m", 0) or 0)
                width_m = float(item.get("width_m", 0) or 0)
                load_class = int(item.get("load_class", 800) or 800)
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
    require_all_priced: bool = True,
    pile_logistics_cost: float = 0,
    pile_trip_overrides: dict[str, int] | None = None,
    pile_catalog_db_path: str | None = None,
) -> dict[str, Any]:
    """
    Рассчитывает общую стоимость заказа.

    unit_price в позициях считается уже с НДС. Скидка применяется к сумме плит.
    НДС для отображения: сумма плит после скидки * VAT_RATE.
    Итого к оплате: сумма плит после скидки + услуга по доставке грузов
    (стоимость рейса × ceil(масса заказа кг / 18600); в базу НДС по плитам не входит).

    subtotal = total_with_vat - vat_amount (согласованная разбивка для документов и архива).
    """
    if require_all_priced:
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
        elif not require_all_priced:
            unit_price = 0.0
        elif _is_pile_item(item):
            mark = str(item.get("mark") or item.get("name") or "").strip()
            grade = str(item.get("concrete_grade") or "B25").strip()
            unit_price = lookup_pile_price(mark, grade, db_path=db_path)
        elif _is_bridge_pile_item(item):
            mark = str(item.get("mark") or item.get("name") or "").strip()
            grade = str(item.get("concrete_grade") or "B25").strip()
            unit_price = lookup_bridge_pile_price(mark, grade, db_path=db_path)
        elif _is_fbs_item(item):
            mark = str(item.get("mark") or item.get("name") or "").strip()
            grade = str(item.get("concrete_grade") or "B25").strip()
            unit_price = lookup_fbs_price(mark, grade, db_path=db_path)
        elif _is_march_item(item):
            mark = str(item.get("mark") or item.get("name") or "").strip()
            grade = str(item.get("concrete_grade") or "B25").strip()
            unit_price = lookup_march_price(mark, grade, db_path=db_path)
        elif _is_step_item(item):
            mark = str(item.get("mark") or item.get("name") or "").strip()
            unit_price = lookup_step_price(mark, db_path=db_path)
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
    cargo_kg = total_order_cargo_weight_kg(order_data, product_types={"plates"})
    plate_delivery_total = delivery_service_charge_rub(trip_cost, cargo_kg)

    pile_trip = max(0.0, float(pile_logistics_cost or 0.0))
    catalog_path = _resolve_pile_catalog_db_path(pile_catalog_db_path, db_path)
    pile_breakdown = _compute_pile_delivery_breakdown(
        order_data,
        overrides=pile_trip_overrides,
        catalog_db_path=catalog_path,
    )
    pile_delivery_total = (
        round(pile_trip * pile_breakdown.total_trips, 2) if pile_breakdown.ready else 0.0
    )
    delivery_total = plate_delivery_total + pile_delivery_total
    vat_amount = round(plates_total_with_vat * VAT_RATE, 2)
    total_with_vat = round(plates_total_with_vat + delivery_total, 2)
    subtotal = round(total_with_vat - vat_amount, 2)

    plate_trips = cargo_delivery_trips_count(cargo_kg)
    return {
        "total_qty": total_qty,
        "subtotal": subtotal,
        "vat_amount": vat_amount,
        "total_with_vat": total_with_vat,
        "plate_delivery_total": plate_delivery_total,
        "pile_delivery_total": pile_delivery_total,
        "pile_trips": pile_breakdown.total_trips,
        "pile_trip_pending_marks": list(pile_breakdown.pending_marks),
        "pile_delivery_ready": pile_breakdown.ready,
        "plate_trips": plate_trips,
    }


def kp_delivery_export_lines(
    totals: Mapping[str, Any],
    *,
    plate_trip_cost: float,
    pile_trip_cost: float,
) -> list[dict[str, Any]]:
    """Строки доставки для PDF/XLSX: плиты и/или сваи, если сумма > 0."""
    plate_amount = float(totals.get("plate_delivery_total") or 0.0)
    pile_amount = float(totals.get("pile_delivery_total") or 0.0)
    pile_ready = bool(totals.get("pile_delivery_ready", True))
    show_pile = pile_ready and pile_amount > 0
    show_plate = plate_amount > 0
    plate_label = "Доставка плит" if show_pile else "Услуга по доставке грузов"
    lines: list[dict[str, Any]] = []
    if show_plate:
        lines.append(
            {
                "label": plate_label,
                "trips": int(totals.get("plate_trips") or 0),
                "unit_price": max(0.0, float(plate_trip_cost or 0.0)),
                "amount": plate_amount,
            }
        )
    if show_pile:
        lines.append(
            {
                "label": "Доставка свай",
                "trips": int(totals.get("pile_trips") or 0),
                "unit_price": max(0.0, float(pile_trip_cost or 0.0)),
                "amount": pile_amount,
            }
        )
    return lines
