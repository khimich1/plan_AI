"""Сборка документа «КоммерческоеПредложение». Файл не пишет."""

from __future__ import annotations

import hashlib
import json
import re
from decimal import Decimal, ROUND_HALF_UP
from typing import Any, Literal, Mapping

INVOICE_DOCUMENT = "КоммерческоеПредложение"
INVOICE_WAREHOUSES = (
    "Склад Готовой Продукции",
    "Склад покупных товаров (для перепродажи)",
    "БСУ склад",
    "Сырье склад производства",
    "Полуфабрикаты склад производства",
)
ORGANIZATION = "ООО «Комбинат ЖБК»"
LINE_UNIT = "шт"
PRICE_KIND = "Продажная"
VAT_RATE_LABEL = "22%"
FILE_STATUS = "в архиве"

InvoiceAction = Literal["create", "update"]

_ISO_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
_RU_DATE = re.compile(r"^(\d{2})\.(\d{2})\.(\d{4})$")
_KOPECK = Decimal("0.01")
_VAT_INCLUDED = Decimal(22) / Decimal(122)


class InvoiceBuildError(ValueError):
    """Документ собрать нельзя: для update нет номера или УИД."""


def build_invoice_document(
    offer: Mapping[str, Any],
    *,
    action: InvoiceAction = "create",
    order_number: str | None = None,
    uid_order: str | None = None,
    uid_kp: str | None = None,
    warehouse: str | None = None,
    line_delivery_kopecks: list[int] | None = None,
) -> tuple[dict[str, Any], str]:
    """Возвращает словарь документа и хеш снимка состава.

    ``create`` не кладёт идентификаторы заказа. ``update`` ставит ``УИД_ЗК``,
    ``Номер_ЗК`` и ``УИД_КП`` первыми ключами и не собирается без любого из них.
    ``Цена`` — цена карточки до скидки. ``СуммаДоставки`` прибавляется после
    скидки. Шапка считается по строкам файла. Склад в хеш не входит.
    НомерДоговора сюда не ставится — его добавляет ``attach_contract_number``.
    """
    if action not in ("create", "update"):
        raise InvoiceBuildError("Неизвестное действие счёта")
    ids: tuple[str, str, str] | None = None
    if action == "update":
        ids = _required_update_ids(order_number, uid_order, uid_kp)

    document = _document(
        offer,
        warehouse=warehouse,
        line_delivery_kopecks=line_delivery_kopecks,
    )
    if ids is not None:
        number, order_uid, kp_uid = ids
        document = {
            "УИД_ЗК": order_uid,
            "Номер_ЗК": number,
            "УИД_КП": kp_uid,
            **document,
        }
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


def _required_update_ids(
    order_number: str | None,
    uid_order: str | None,
    uid_kp: str | None,
) -> tuple[str, str, str]:
    number = (order_number or "").strip()
    order_uid = (uid_order or "").strip()
    kp_uid = (uid_kp or "").strip()
    if not number or not order_uid or not kp_uid:
        raise InvoiceBuildError("Исправление без номера заказа или УИД не собирается")
    return number, order_uid, kp_uid


def _document(
    offer: Mapping[str, Any],
    *,
    warehouse: str | None,
    line_delivery_kopecks: list[int] | None,
) -> dict[str, Any]:
    counterparty = offer.get("counterparty") or {}
    if not isinstance(counterparty, Mapping):
        counterparty = {}
    lines = [line for line in offer.get("lines") or [] if isinstance(line, Mapping)]
    shares = _delivery_shares(line_delivery_kopecks, len(lines))
    goods, totals = _goods_and_totals(offer, lines, shares)
    manager = offer.get("manager") or offer.get("manager_name")
    document: dict[str, Any] = {
        "Документ": INVOICE_DOCUMENT,
        "НомерВПриложении": offer.get("kp_id"),
        "Дата": _invoice_date(offer.get("date") or offer.get("creation_date")),
        "Организация": ORGANIZATION,
        "Клиент": counterparty.get("name"),
        "Контрагент": {
            "Наименование": counterparty.get("name"),
            "Код": counterparty.get("code"),
            "GUID": counterparty.get("guid"),
            "ИНН": counterparty.get("inn"),
            "КПП": counterparty.get("kpp"),
        },
        "КонтактноеЛицо": None,
        "Валюта": "RUB",
        "ХозяйственнаяОперация": "РеализацияКлиенту",
        "Налогообложение": "ОблагаетсяНДС",
        "ЦенаВключаетНДС": True,
        "Статус": FILE_STATUS,
        "Комментарий": "",
        "Менеджер": manager,
        "Автор": manager,
    }
    warehouse_name = str(warehouse or "").strip()
    if warehouse_name:
        document["Склад"] = warehouse_name
    document["Самовывоз"] = all(kopecks <= 0 for kopecks in shares)
    document["ПроцентСкидки"] = offer.get("discount_percent")
    document["СуммаСкидки"] = _json_rub(totals["discount"])
    document["СуммаНДС"] = _json_rub(totals["vat"])
    document["СуммаДокумента"] = _json_rub(totals["payable"])
    document["Товары"] = goods
    return document


def _invoice_date(raw: Any) -> str | None:
    text = str(raw).strip() if raw is not None else ""
    if not text:
        return None
    if _ISO_DATE.fullmatch(text):
        return text
    matched = _RU_DATE.fullmatch(text)
    if matched is None:
        return text
    day, month, year = matched.groups()
    return f"{year}-{month}-{day}"


def _delivery_shares(raw: list[int] | None, count: int) -> list[int]:
    if raw is None:
        return [0] * count
    if len(raw) != count:
        raise InvoiceBuildError("Доля доставки не совпадает с числом строк")
    return [int(kopecks) for kopecks in raw]


def _goods_and_totals(
    offer: Mapping[str, Any],
    lines: list[Mapping[str, Any]],
    shares: list[int],
) -> tuple[list[dict[str, Any]], dict[str, Decimal]]:
    goods: list[dict[str, Any]] = []
    payable = Decimal(0)
    vat = Decimal(0)
    discount = Decimal(0)
    for line, kopecks in zip(lines, shares):
        money = _line_money(offer, line, kopecks)
        payable += money["payable"]
        vat += money["vat"]
        discount += money["discount"]
        goods.append(_goods_line(offer, line, money))
    return goods, {"payable": payable, "vat": vat, "discount": discount}


def _line_money(
    offer: Mapping[str, Any],
    line: Mapping[str, Any],
    delivery_kopecks: int,
) -> dict[str, Decimal]:
    payable = _line_payable(offer, line, delivery_kopecks)
    vat = (payable * _VAT_INCLUDED).quantize(_KOPECK, rounding=ROUND_HALF_UP)
    return {
        "payable": payable,
        "vat": vat,
        "net": payable - vat,
        "discount": _line_discount_amount(offer, line),
        "delivery": _kopecks_to_rub(delivery_kopecks),
    }


def _line_payable(offer: Mapping[str, Any], line: Mapping[str, Any], delivery_kopecks: int) -> Decimal:
    price = line.get("price") if "price" in line else 0
    product = _decimal(price) * _discount_factor(_line_discount(offer, line)) * _qty(line.get("qty"))
    return (product + _kopecks_to_rub(delivery_kopecks)).quantize(_KOPECK, rounding=ROUND_HALF_UP)


def _line_discount_amount(offer: Mapping[str, Any], line: Mapping[str, Any]) -> Decimal:
    percent = _clamped_percent(_line_discount(offer, line))
    price = line.get("price") if "price" in line else 0
    amount = _decimal(price) * Decimal(_qty(line.get("qty"))) * percent / Decimal(100)
    return amount.quantize(_KOPECK, rounding=ROUND_HALF_UP)


def _line_discount(offer: Mapping[str, Any], line: Mapping[str, Any]) -> Any:
    if "discount_percent" in line:
        return line.get("discount_percent")
    return offer.get("discount_percent")


def _kopecks_to_rub(delivery_kopecks: int) -> Decimal:
    return Decimal(int(delivery_kopecks)) / Decimal(100)


def _discount_factor(discount: Any) -> Decimal:
    return Decimal(1) - _clamped_percent(discount) / Decimal(100)


def _clamped_percent(discount: Any) -> Decimal:
    percent = _decimal(discount)
    if percent < 0:
        return Decimal(0)
    if percent > 100:
        return Decimal(100)
    return percent


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


def _goods_line(
    offer: Mapping[str, Any],
    line: Mapping[str, Any],
    money: Mapping[str, Decimal],
) -> dict[str, Any]:
    price = line.get("price") if "price" in line else 0
    return {
        "Номенклатура": line.get("name") or line.get("nomenclature") or "",
        "GUID": line.get("guid"),
        "Характеристика": None,
        "Количество": line.get("qty"),
        "ЕдиницаИзмерения": LINE_UNIT,
        "ВидЦены": PRICE_KIND,
        "Цена": price,
        "ПроцентСкидки": _line_discount(offer, line),
        "СуммаСкидки": _json_rub(money["discount"]),
        "Сумма": _json_rub(money["net"]),
        "СтавкаНДС": VAT_RATE_LABEL,
        "СуммаНДС": _json_rub(money["vat"]),
        "СуммаСНДС": _json_rub(money["payable"]),
        "СуммаДоставки": _json_rub(money["delivery"]),
    }
