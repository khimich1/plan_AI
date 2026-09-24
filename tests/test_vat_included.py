"""НДС «в том числе 22%» по строкам: эталон счёта 680 и порог старого kp_id."""

from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP

from core.commercial_pricing import (
    VAT_PER_LINE_FROM_KP_ID,
    calculate_total_cost,
    vat_included_from_line_sums,
)

# Суммы строк счёта 680, рубли. Итог 2 479 416,00.
INVOICE_680_LINE_SUMS = (
    "228062.15",
    "42931.57",
    "483560.67",
    "41901.07",
    "41901.07",
    "219500.38",
    "22924.21",
    "135295.48",
    "378473.67",
    "270120.55",
    "19132.61",
    "235664.23",
    "18709.40",
    "56128.21",
    "88374.08",
    "80255.98",
    "17139.27",
    "30154.99",
    "1609.46",
    "34138.53",
    "33438.42",
)


def test_invoice_680_vat_is_sum_of_per_line_rounds() -> None:
    lines = [Decimal(amount) for amount in INVOICE_680_LINE_SUMS]
    result = vat_included_from_line_sums(lines)

    grand = sum(lines, Decimal("0"))
    single_round = (grand * Decimal(22) / Decimal(122)).quantize(
        Decimal("0.01"), rounding=ROUND_HALF_UP
    )

    assert grand == Decimal("2479416.00")
    assert single_round == Decimal("447107.80")
    assert result == Decimal("447107.82")
    assert result != single_round


def test_vat_included_zero_line_and_empty_list() -> None:
    assert vat_included_from_line_sums([]) == Decimal("0.00")
    assert vat_included_from_line_sums([Decimal("0")]) == Decimal("0.00")


def _priced_plate(*, unit_price: float, qty: int = 1, weight_kg: float = 1000.0) -> dict:
    return {
        "name": "ПБ",
        "length_m": 6.0,
        "width_m": 1.2,
        "qty": qty,
        "load_class": 800,
        "unit_price": unit_price,
        "weight": weight_kg,
    }


def test_old_kp_id_keeps_plates_times_022_without_delivery() -> None:
    """kp_id ниже порога: НДС только с плит после скидки × 0,22, доставка вне базы."""
    order = [_priced_plate(unit_price=1000.0)]
    assert VAT_PER_LINE_FROM_KP_ID == 27
    old = calculate_total_cost(
        order,
        discount_percent=0,
        logistics_cost=500.0,
        db_path=":memory:",
        require_all_priced=False,
        kp_id=26,
    )
    assert old["vat_amount"] == round(1000.0 * 0.22, 2)
    assert old["vat_amount"] == 220.0
    assert old["total_with_vat"] == round(1000.0 + 500.0, 2)
    assert old["vat_amount"] != 420584.26


def test_new_kp_id_uses_per_line_vat_and_keeps_payable_total() -> None:
    order = [_priced_plate(unit_price=1000.0, qty=2)]
    common = dict(
        discount_percent=0,
        logistics_cost=500.0,
        db_path=":memory:",
        require_all_priced=False,
    )
    old = calculate_total_cost(order, kp_id=26, **common)
    new = calculate_total_cost(order, kp_id=VAT_PER_LINE_FROM_KP_ID, **common)
    preview = calculate_total_cost(order, kp_id=None, **common)

    product_line = Decimal("2000.00")
    delivery_line = Decimal(str(old["plate_delivery_total"]))
    expected_vat = vat_included_from_line_sums([product_line, delivery_line])

    assert new["total_with_vat"] == old["total_with_vat"]
    assert new["vat_amount"] == float(expected_vat)
    assert new["subtotal"] == round(new["total_with_vat"] - new["vat_amount"], 2)
    assert preview["vat_amount"] == new["vat_amount"]
    assert preview["total_with_vat"] == old["total_with_vat"]
    assert new["vat_amount"] != old["vat_amount"]
