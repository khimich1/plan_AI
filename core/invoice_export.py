"""Сборка документа «Счёт на оплату». Файл не пишет."""

from __future__ import annotations

import hashlib
import json
from decimal import Decimal, ROUND_HALF_UP
from typing import Any, Literal, Mapping

from core.commercial_pricing import vat_included_from_line_sums

INVOICE_DOCUMENT = "Счёт на оплату"
INVOICE_EVENT = "invoice_export"
INVOICE_WAREHOUSES = (
    "Склад Готовой Продукции",
    "Склад покупных товаров (для перепродажи)",
    "БСУ склад",
    "Сырье склад производства",
    "Полуфабрикаты склад производства",
)

InvoiceAction = Literal["create", "update"]


class InvoiceBuildError(ValueError):
    """Документ собрать нельзя: для update нет номера счёта."""


_KOPECK = Decimal("0.01")


def build_invoice_document(
    offer: Mapping[str, Any],
    *,
    action: InvoiceAction = "create",
    order_number: str | None = None,
    warehouse: str | None = None,
    line_delivery_kopecks: list[int] | None = None,
) -> tuple[dict[str, Any], str]:
    """Возвращает словарь документа и хеш снимка состава.

    ``Цена`` — цена карточки до скидки. ``СуммаДоставки`` прибавляется после
    скидки. Шапка считается по строкам файла. Склад в хеш не входит.
    НомерДоговора сюда не ставится — его добавляет ``attach_contract_number``.
    """
    if action not in ("create", "update"):
        raise InvoiceBuildError("Неизвестное действие счёта")
    number = (order_number or "").strip()
    if action == "update" and not number:
        raise InvoiceBuildError("Исправление без номера счёта не собирается")

    document = _document(
        offer,
        action=action,
        order_number=number if action == "update" else None,
        warehouse=warehouse,
        line_delivery_kopecks=line_delivery_kopecks,
    )
    return document, invoice_snapshot_hash(offer)


def invoice_snapshot_hash(offer: Mapping[str, Any]) -> str:
    """Хеш состава: строки, скидка заказа, доставка, итог."""
    lines = []
    for line in offer.get("lines") or []:
        if not isinstance(line, Mapping):
            continue
        discount = line.get("discount") if "discount" in line else line.get("discount_percent")
        lines.append(
            {
                "nomenclature": line.get("name") or line.get("nomenclature") or "",
                "qty": line.get("qty"),
                "price": line.get("price"),
                "discount": discount,
            }
        )
    payload = {
        "lines": lines,
        "discount": offer.get("discount_percent"),
        "delivery": offer.get("delivery"),
        "total": offer.get("total_amount"),
    }
    raw = json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def card_offer_from_kp(raw: Mapping[str, Any]) -> dict[str, Any]:
    """Цифры карточки этого kp_id. Отдельную формулу НДС не считает."""
    discount = raw.get("discount_percent")
    lines: list[dict[str, Any]] = []
    for plate in raw.get("plates") or []:
        lines.append(_stored_line(plate, "plate_name", "plate", discount))
    for key, kind in (
        ("piles", "pile"),
        ("steps", "stair_step"),
        ("marches", "stair_flight"),
        ("bridge_piles", "bridge_pile"),
        ("fbs", "fbs"),
    ):
        for item in raw.get(key) or []:
            lines.append(_stored_line(item, "mark", kind, discount))
    return {
        "kp_id": raw.get("kp_id"),
        "date": raw.get("creation_date"),
        "manager": raw.get("manager_name"),
        "discount_percent": discount,
        "vat_amount": raw.get("vat_amount"),
        "total_amount": raw.get("total_amount"),
        "delivery": raw.get("logistics_cost"),
        "counterparty": {
            "name": raw.get("customer_name"),
            "code": None,
            "guid": None,
            "inn": raw.get("customer_inn"),
            "kpp": raw.get("customer_kpp"),
        },
        "lines": lines,
    }


def _stored_line(
    item: Mapping[str, Any],
    name_key: str,
    product_kind: str,
    order_discount: Any,
) -> dict[str, Any]:
    line: dict[str, Any] = {
        "name": item.get(name_key) or "",
        "qty": item.get("qty"),
        "product_kind": product_kind,
    }
    if item.get("unit_price") is not None:
        line["price"] = item.get("unit_price")
    if order_discount is not None:
        line["discount_percent"] = order_discount
    return line


def _document(
    offer: Mapping[str, Any],
    *,
    action: InvoiceAction,
    order_number: str | None,
    warehouse: str | None,
    line_delivery_kopecks: list[int] | None,
) -> dict[str, Any]:
    counterparty = offer.get("counterparty") or {}
    if not isinstance(counterparty, Mapping):
        counterparty = {}
    lines = [line for line in offer.get("lines") or [] if isinstance(line, Mapping)]
    shares = _delivery_shares(line_delivery_kopecks, len(lines))
    goods, payables = _goods_and_payables(offer, lines, shares)
    document: dict[str, Any] = {
        "event": INVOICE_EVENT,
        "action": action,
        "Документ": INVOICE_DOCUMENT,
        "НомерВПриложении": offer.get("kp_id"),
    }
    if order_number:
        document["НомерЗаказа"] = order_number
    document["Дата"] = offer.get("date") or offer.get("creation_date")
    document["Контрагент"] = {
        "Наименование": counterparty.get("name"),
        "Код": counterparty.get("code"),
        "GUID": counterparty.get("guid"),
        "ИНН": counterparty.get("inn"),
        "КПП": counterparty.get("kpp"),
    }
    warehouse_name = str(warehouse or "").strip()
    if warehouse_name:
        document["Склад"] = warehouse_name
    document["Самовывоз"] = all(kopecks <= 0 for kopecks in shares)
    document["Менеджер"] = offer.get("manager") or offer.get("manager_name")
    if "discount_percent" in offer:
        document["ПроцентСкидки"] = offer.get("discount_percent")
    if "discount_amount" in offer:
        document["СуммаСкидки"] = offer.get("discount_amount")
    document["СуммаНДС"] = _json_rub(vat_included_from_line_sums(payables))
    document["СуммаДокумента"] = _json_rub(sum(payables, Decimal("0")))
    document["Товары"] = goods
    return document


def _delivery_shares(raw: list[int] | None, count: int) -> list[int]:
    if raw is None:
        return [0] * count
    if len(raw) != count:
        raise InvoiceBuildError("Доля доставки не совпадает с числом строк")
    return [int(kopecks) for kopecks in raw]


def _goods_and_payables(
    offer: Mapping[str, Any],
    lines: list[Mapping[str, Any]],
    shares: list[int],
) -> tuple[list[dict[str, Any]], list[Decimal]]:
    goods: list[dict[str, Any]] = []
    payables: list[Decimal] = []
    for line, kopecks in zip(lines, shares):
        payables.append(_line_payable(offer, line, kopecks))
        goods.append(_goods_line(line, delivery_kopecks=kopecks))
    return goods, payables


def _line_payable(offer: Mapping[str, Any], line: Mapping[str, Any], delivery_kopecks: int) -> Decimal:
    price = line.get("price") if "price" in line else 0
    discount = line.get("discount_percent") if "discount_percent" in line else offer.get("discount_percent")
    product = _decimal(price) * _discount_factor(discount) * _qty(line.get("qty"))
    delivery = Decimal(int(delivery_kopecks)) / Decimal(100)
    return (product + delivery).quantize(_KOPECK, rounding=ROUND_HALF_UP)


def _discount_factor(discount: Any) -> Decimal:
    percent = _decimal(discount)
    if percent < 0:
        percent = Decimal(0)
    if percent > 100:
        percent = Decimal(100)
    return Decimal(1) - percent / Decimal(100)


def _qty(raw: Any) -> int:
    try:
        return max(0, int(raw or 0))
    except (TypeError, ValueError):
        return 0


def _decimal(raw: Any) -> Decimal:
    if raw is None or raw == "":
        return Decimal(0)
    return Decimal(str(raw))


def _json_rub(amount: Decimal) -> float:
    return float(format(amount.quantize(_KOPECK, rounding=ROUND_HALF_UP), "f"))


def _goods_line(line: Mapping[str, Any], *, delivery_kopecks: int) -> dict[str, Any]:
    item: dict[str, Any] = {
        "Номенклатура": line.get("name") or line.get("nomenclature") or "",
        "Количество": line.get("qty"),
    }
    if "guid" in line:
        item["GUID"] = line.get("guid")
    if "price" in line:
        item["Цена"] = line.get("price")
    if "discount_percent" in line:
        item["ПроцентСкидки"] = line.get("discount_percent")
    item["СуммаДоставки"] = _json_rub(Decimal(int(delivery_kopecks)) / Decimal(100))
    if "discount_amount" in line:
        item["СуммаСкидки"] = line.get("discount_amount")
    if "sum" in line:
        item["Сумма"] = line.get("sum")
    if "vat_amount" in line:
        item["СуммаНДС"] = line.get("vat_amount")
    if "sum_with_vat" in line:
        item["СуммаСНДС"] = line.get("sum_with_vat")
    return item
