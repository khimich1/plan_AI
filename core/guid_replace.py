"""Сборка строки замены марки и окно «было / станет».

Имя плиты — то, что выбрал человек: без повторного LIKE и canonical_name.
Нагрузка плиты — load_class_from_plate_name (12,5 → 1250), не parse_load_code_from_name.
"""

from __future__ import annotations

import copy
import re
from dataclasses import dataclass
from typing import Any, Mapping

from core.commercial_pricing import calculate_total_cost
from core.config_and_data import parse_name_to_sizes
from core.kp_db_nomenclature import load_class_from_plate_name
from core.kp_plate_weight import resolve_kp_line_weight_kg

_ONEOFF = "oneoff"
_ONEOFF_NOTICE = "Договорная цена будет заменена ценой прайса."
_LENGTH_DM_RE = re.compile(r"П[БК]\s+([\d,]+)\s*-", re.IGNORECASE)
_LOAD_TOKEN_RE = re.compile(r"\d+(?:[.,]\d+)?п", re.IGNORECASE)

_LINE_TYPE = {
    "plate": "plates",
    "plates": "plates",
    "pile": "piles",
    "piles": "piles",
    "bridge_pile": "bridge_piles",
    "bridge_piles": "bridge_piles",
    "composite_pile": "composite_piles",
    "composite_piles": "composite_piles",
    "fbs": "fbs",
    "stair_step": "steps",
    "step": "steps",
    "steps": "steps",
    "stair_flight": "marches",
    "march": "marches",
    "marches": "marches",
}


class GuidReplaceError(ValueError):
    """Отказ замены. ``detail`` — русская фраза для ответа API."""

    def __init__(self, detail: str) -> None:
        self.detail = detail
        super().__init__(detail)


@dataclass(frozen=True)
class ReplaceTarget:
    mark: str
    nomenclature_id: str | None = None


@dataclass(frozen=True)
class MoneyView:
    unit_price: float
    unit_discounted: float
    line_sum: float
    total: float
    subtotal: float
    vat_amount: float


@dataclass(frozen=True)
class ReplacementPreview:
    qty: int
    was: MoneyView
    will: MoneyView
    delivery_was: float | None
    delivery_will: float | None
    oneoff_notice: str | None


def build_replacement(
    line: Mapping[str, Any],
    target: ReplaceTarget,
    discount_percent: float,
    kind: str,
    *,
    actor: str,
    pb_db_path: str,
    replace_oneoff: bool = False,
) -> dict[str, Any]:
    """Новая строка заказа: выбранная марка, цена прайса и скидка этого КП.

    Количество, line_id, статус, plan_id и класс бетона не меняются.
    """
    chosen = str(target.mark or "").strip()
    if not chosen:
        raise GuidReplaceError("Не выбрана марка для замены.")
    line_type = _LINE_TYPE.get(str(kind or "").strip().lower())
    if line_type is None:
        raise GuidReplaceError(f"Неизвестный вид изделия «{kind}».")
    _guard_oneoff(line, actor, replace_oneoff)
    new = dict(line)
    factor = _discount_factor(discount_percent)
    if line_type == "plates":
        _apply_plate(new, chosen, target.nomenclature_id, factor, pb_db_path)
    else:
        grade = str(line.get("concrete_grade") or "").strip()
        price = _price_for_kind(line_type, chosen, grade, pb_db_path)
        if price is None:
            if line_type == "steps":
                raise GuidReplaceError(f"Нет цены ступени «{chosen}» в прайсе.")
            raise GuidReplaceError(
                f"Нет цены марки «{chosen}» в классе бетона {grade or '—'}."
            )
        new["mark"] = chosen
        new["name"] = chosen
        new["unit_price"] = float(price)
        new["discounted_price"] = float(price) * factor
    if str(line.get("price_source") or "").strip() == _ONEOFF:
        new.pop("price_source", None)
    return new


def preview_replacement(
    order_data: list[dict[str, Any]],
    *,
    source_mark: str,
    target: ReplaceTarget,
    discount_percent: float,
    kind: str,
    actor: str,
    pb_db_path: str,
    totals_db_path: str,
    kp_id: int | None = None,
    logistics_cost: float = 0,
) -> ReplacementPreview:
    """Деньги одного КП до и после замены всех строк этой марки.

    Строка доставки есть в ответе только если её сумма изменилась.
    Договорная строка у экономиста считается по прайсу и помечается фразой.
    """
    needle = str(source_mark or "").strip()
    indexes = [
        index
        for index, line in enumerate(order_data)
        if _line_mark(line, kind) == needle
    ]
    if not indexes:
        raise GuidReplaceError(f"В составе нет марки «{needle}».")
    oneoff = any(
        str(order_data[index].get("price_source") or "").strip() == _ONEOFF
        for index in indexes
    )
    who = str(actor or "").strip().casefold()
    if oneoff and who != "economist":
        raise GuidReplaceError("Договорную цену менеджер заменить не может.")
    notice = _ONEOFF_NOTICE if oneoff and who == "economist" else None
    new_order = [dict(line) for line in order_data]
    built: list[dict[str, Any]] = []
    for index in indexes:
        row = build_replacement(
            order_data[index],
            target,
            discount_percent,
            kind,
            actor=actor,
            pb_db_path=pb_db_path,
            replace_oneoff=oneoff and who == "economist",
        )
        new_order[index] = row
        built.append(row)
    before = _totals(order_data, discount_percent, totals_db_path, kp_id, logistics_cost)
    after = _totals(new_order, discount_percent, totals_db_path, kp_id, logistics_cost)
    delivery_before = _delivery_sum(before)
    delivery_after = _delivery_sum(after)
    show_delivery = abs(delivery_before - delivery_after) >= 0.01
    originals = [order_data[index] for index in indexes]
    qty = sum(int(row.get("qty") or 0) for row in originals)
    return ReplacementPreview(
        qty=qty,
        was=_money_view(originals, discount_percent, before),
        will=_money_view(built, discount_percent, after),
        delivery_was=delivery_before if show_delivery else None,
        delivery_will=delivery_after if show_delivery else None,
        oneoff_notice=notice,
    )


def _guard_oneoff(line: Mapping[str, Any], actor: str, replace_oneoff: bool) -> None:
    if str(line.get("price_source") or "").strip() != _ONEOFF:
        return
    if str(actor or "").strip().casefold() != "economist":
        raise GuidReplaceError("Договорную цену менеджер заменить не может.")
    if not replace_oneoff:
        raise GuidReplaceError(
            "Договорная цена будет заменена ценой прайса. Подтвердите замену."
        )


def _apply_plate(
    new: dict[str, Any],
    chosen: str,
    nomenclature_id: str | None,
    factor: float,
    pb_db_path: str,
) -> None:
    from core.commercial_pricing import PriceNotFoundError, lookup_plate_price

    length_m, width_m = parse_name_to_sizes(chosen)
    if length_m is None or width_m is None:
        raise GuidReplaceError(f"Не удалось прочитать размеры плиты «{chosen}».")
    load_class = _plate_load_class(chosen)
    if load_class is None:
        raise GuidReplaceError(f"Не удалось прочитать нагрузку из имени «{chosen}».")
    try:
        unit = lookup_plate_price(length_m, width_m, load_class, db_path=pb_db_path)
    except PriceNotFoundError as exc:
        raise GuidReplaceError(f"Нет цены плиты «{chosen}» в прайсе.") from exc
    qty = int(new.get("qty") or 0)
    _unit_weight, total_weight = resolve_kp_line_weight_kg(
        {"length_m": length_m, "width_m": width_m, "qty": qty}
    )
    length_match = _LENGTH_DM_RE.search(chosen)
    new["name"] = chosen
    new["product_type"] = "plates"
    new["length_m"] = length_m
    new["width_m"] = width_m
    new["load_class"] = load_class
    new["length_dm_raw"] = (
        length_match.group(1).strip() if length_match else str(new.get("length_dm_raw") or "")
    )
    new["weight"] = total_weight
    new["nomenclature_id"] = nomenclature_id
    new["unit_price"] = float(unit)
    new["discounted_price"] = float(unit) * factor


def _plate_load_class(name: str) -> int | None:
    """Нагрузка карточки. У особой «…12п Пв-7.1» хвост срезается и читается тем же хелпером."""
    found = load_class_from_plate_name(name)
    if found is not None:
        return found
    token = _LOAD_TOKEN_RE.search(name)
    if token is None:
        return None
    return load_class_from_plate_name(name[: token.end()])


def _price_for_kind(
    line_type: str,
    mark: str,
    concrete_grade: str,
    pb_db_path: str,
) -> float | None:
    grade = concrete_grade or "B25"
    if line_type == "piles":
        from core.pile_price_db import get_pile_price

        return get_pile_price(mark, grade, pb_db_path)
    if line_type == "bridge_piles":
        from core.bridge_pile_price_db import get_bridge_pile_price

        return get_bridge_pile_price(mark, grade, pb_db_path)
    if line_type == "composite_piles":
        from core.composite_pile_price_db import get_composite_pile_price

        return get_composite_pile_price(mark, grade, pb_db_path)
    if line_type == "fbs":
        from core.fbs_price_db import get_fbs_price

        return get_fbs_price(mark, grade, pb_db_path)
    if line_type == "marches":
        from core.march_price_db import get_march_price

        return get_march_price(mark, grade, pb_db_path)
    if line_type == "steps":
        from core.step_price_db import get_step_price

        return get_step_price(mark, pb_db_path)
    return None


def _line_mark(line: Mapping[str, Any], kind: str) -> str:
    line_type = _LINE_TYPE.get(str(kind or "").strip().lower())
    if line_type == "plates":
        return str(line.get("name") or "").strip()
    return str(line.get("mark") or line.get("name") or "").strip()


def _discount_factor(discount_percent: float) -> float:
    percent = min(max(float(discount_percent or 0.0), 0.0), 100.0)
    return 1.0 - percent / 100.0


def _totals(
    order_data: list[dict[str, Any]],
    discount_percent: float,
    totals_db_path: str,
    kp_id: int | None,
    logistics_cost: float,
) -> dict[str, Any]:
    return calculate_total_cost(
        copy.deepcopy(order_data),
        discount_percent,
        logistics_cost=logistics_cost,
        db_path=totals_db_path,
        require_all_priced=False,
        pile_catalog_db_path=totals_db_path,
        weight_catalog_db_path=totals_db_path,
        kp_id=kp_id,
    )


def _delivery_sum(totals: Mapping[str, Any]) -> float:
    return float(
        float(totals.get("plate_delivery_total") or 0.0)
        + float(totals.get("pile_delivery_total") or 0.0)
        + float(totals.get("fbs_lm_delivery_total") or 0.0)
        + float(totals.get("long_pile_delivery_total") or 0.0)
    )


def _money_view(
    lines: list[Mapping[str, Any]],
    discount_percent: float,
    totals: Mapping[str, Any],
) -> MoneyView:
    factor = _discount_factor(discount_percent)
    first = lines[0]
    unit = float(first.get("unit_price") or 0.0)
    stored = first.get("discounted_price")
    discounted = float(stored) if stored is not None else unit * factor
    line_sum = 0.0
    for line in lines:
        qty = int(line.get("qty") or 0)
        row_unit = float(line.get("unit_price") or 0.0)
        row_stored = line.get("discounted_price")
        row_discounted = float(row_stored) if row_stored is not None else row_unit * factor
        line_sum += row_discounted * qty
    return MoneyView(
        unit_price=unit,
        unit_discounted=discounted,
        line_sum=line_sum,
        total=float(totals["total_with_vat"]),
        subtotal=float(totals["subtotal"]),
        vat_amount=float(totals["vat_amount"]),
    )
