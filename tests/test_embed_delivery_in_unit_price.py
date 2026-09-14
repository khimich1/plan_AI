"""RED/GREEN: embed delivery into unit prices after product discount.

Pure math lives in ``core.embed_delivery_in_unit_price``. XLSX flag tests
live in the same module after the helper is green.
"""

from __future__ import annotations

import inspect
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from core.commercial_offer_xlsx import calculate_total_cost, generate_commercial_offer_xlsx
from core.commercial_pricing import VAT_RATE
from core.embed_delivery_in_unit_price import embed_delivery_in_unit_prices


def _embed(**kwargs):
    return embed_delivery_in_unit_prices(**kwargs)


def test_discount_applies_only_to_product_not_delivery() -> None:
    """50% off 122 + delivery 100 on qty 1 → 61 + 100 = 161, not 50% of 222."""
    result = _embed(
        qty_by_index=[1],
        product_type_by_index=["plates"],
        unit_price_by_index=[122.0],
        discount_percent=50.0,
        plate_delivery_total=100.0,
        pile_delivery_total=0.0,
    )

    assert result.embedded_plate is True
    assert result.embedded_pile is False
    assert len(result.lines) == 1
    line = result.lines[0]
    assert line.qty == 1
    assert line.line_sum == 161.0
    assert line.surcharge_unit == 100.0
    assert line.discounted_unit == 61.0
    assert line.unit_price == 161.0


def test_mixed_pools_do_not_cross() -> None:
    """Plate delivery only on plates qty; pile delivery only on piles/bridge_piles."""
    result = _embed(
        qty_by_index=[2, 3, 4, 5],
        product_type_by_index=["plates", "piles", "bridge_piles", "fbs"],
        unit_price_by_index=[100.0, 200.0, 300.0, 400.0],
        discount_percent=0.0,
        plate_delivery_total=20.0,
        pile_delivery_total=70.0,
    )

    assert result.embedded_plate is True
    assert result.embedded_pile is True
    plate, pile, bridge, fbs = result.lines
    assert plate.line_sum == 2 * 100.0 + 20.0
    assert pile.surcharge_unit * pile.qty + bridge.surcharge_unit * bridge.qty == 70.0
    assert abs(
        (pile.line_sum - 3 * 200.0) + (bridge.line_sum - 4 * 300.0) - 70.0
    ) < 0.01
    assert fbs.surcharge_unit == 0.0
    assert fbs.line_sum == 5 * 400.0
    surcharge_total = sum(line.surcharge_unit * line.qty for line in result.lines)
    assert abs(surcharge_total - 90.0) < 0.01


def test_divmod_extra_kopecks_sum_to_delivery() -> None:
    """100.01 ₽ (10001 kop) over qty 3 distributes extras in line order."""
    result = _embed(
        qty_by_index=[1, 1, 1],
        product_type_by_index=["plates", "plates", "plates"],
        unit_price_by_index=[10.0, 10.0, 10.0],
        discount_percent=0.0,
        plate_delivery_total=100.01,
        pile_delivery_total=0.0,
    )

    delivery_kop = [round(line.line_sum * 100) - 10 * 100 for line in result.lines]
    assert sum(delivery_kop) == 10001
    assert delivery_kop == [3334, 3334, 3333]


def test_delivery_zero_means_zero_surcharge() -> None:
    result = _embed(
        qty_by_index=[2],
        product_type_by_index=["plates"],
        unit_price_by_index=[122.0],
        discount_percent=50.0,
        plate_delivery_total=0.0,
        pile_delivery_total=0.0,
    )

    assert result.embedded_plate is False
    assert result.embedded_pile is False
    line = result.lines[0]
    assert line.surcharge_unit == 0.0
    assert line.line_sum == 122.0
    assert line.discounted_unit == 61.0


def test_plate_delivery_with_zero_plate_qty_is_not_embedded() -> None:
    result = _embed(
        qty_by_index=[0, 2],
        product_type_by_index=["plates", "piles"],
        unit_price_by_index=[100.0, 200.0],
        discount_percent=0.0,
        plate_delivery_total=50.0,
        pile_delivery_total=0.0,
    )

    assert result.embedded_plate is False
    assert result.embedded_pile is False
    assert result.lines[0].surcharge_unit == 0.0
    assert result.lines[1].surcharge_unit == 0.0
    assert result.lines[1].line_sum == 400.0


def test_qty_zero_line_is_skipped_for_pool_qty() -> None:
    result = _embed(
        qty_by_index=[0, 2],
        product_type_by_index=["plates", "plates"],
        unit_price_by_index=[100.0, 50.0],
        discount_percent=0.0,
        plate_delivery_total=10.0,
        pile_delivery_total=0.0,
    )

    assert result.embedded_plate is True
    assert result.lines[0].line_sum == 0.0
    assert result.lines[1].line_sum == 2 * 50.0 + 10.0


def test_legacy_missing_type_counts_as_plates() -> None:
    result = _embed(
        qty_by_index=[1],
        product_type_by_index=[""],
        unit_price_by_index=[100.0],
        discount_percent=0.0,
        plate_delivery_total=40.0,
        pile_delivery_total=0.0,
    )

    assert result.embedded_plate is True
    assert result.lines[0].line_sum == 140.0


# --- Task 3/4: XLSX generator flag -------------------------------------------


def _discount_plate_order() -> list[dict[str, Any]]:
    return [
        {
            "name": "ПБ 59-12-8п",
            "qty": 1,
            "unit_price": 122.0,
            "length_m": 1.0,
            "width_m": 1.0,
            "product_type": "plates",
        }
    ]


def _mixed_order() -> list[dict[str, Any]]:
    return [
        {
            "line_id": "p1",
            "product_type": "plates",
            "name": "ПБ 60-12-8п",
            "mark": "ПБ 60-12-8п",
            "length_m": 1.0,
            "width_m": 1.0,
            "load_class": 800,
            "qty": 65,
            "unit_price": 1000.0,
            "weight": 500.0,
            "concrete_grade": "М500",
        },
        {
            "line_id": "s1",
            "product_type": "piles",
            "product_kind": "pile",
            "name": "С60.30",
            "mark": "С60.30",
            "concrete_grade": "B25",
            "qty": 14,
            "unit_price": 50.0,
        },
    ]


def _seed_pile_catalog(tmp_path: Path) -> str:
    from core.kp_db_schema import init_schema
    from core.pile_catalog import PileCatalogEntry, upsert_pile_catalog

    db_path = str(tmp_path / "plita.db")
    init_schema(db_path)
    upsert_pile_catalog(
        db_path,
        [PileCatalogEntry("С60.30", 6.0, 300, 0.55, 1380.0, 14)],
    )
    return db_path


def _xlsx_kwargs(order: list[dict[str, Any]], **overrides: Any) -> dict[str, Any]:
    kwargs: dict[str, Any] = {
        "order_data": order,
        "offer_number": "E1",
        "offer_date": "10.09.2026",
        "customer_name": "ООО Тест",
        "kp_db_id": 1,
        "discount_percent": 0.0,
        "logistics_cost": 0.0,
    }
    kwargs.update(overrides)
    return kwargs


def _load_kp_sheet(order: list[dict[str, Any]], **overrides: Any):
    buf = generate_commercial_offer_xlsx(**_xlsx_kwargs(order, **overrides))
    return load_workbook(buf)["КП"]


def _table_layout(ws) -> tuple[int, int, int, int]:
    """Return header_row, name_col, sum_col, first_data_row (1-based)."""
    for row in ws.iter_rows(min_row=1, max_row=40, max_col=8):
        values = [cell.value for cell in row]
        if values and values[0] == "№" and "Наименование" in values and "Сумма" in values:
            header_row = row[0].row
            name_col = values.index("Наименование") + 1
            sum_col = values.index("Сумма") + 1
            return header_row, name_col, sum_col, header_row + 1
    raise AssertionError("table header with Наименование/Сумма not found")


def _product_and_delivery_rows(ws) -> tuple[list[tuple[str, Any]], list[str]]:
    header_row, name_col, sum_col, first_data = _table_layout(ws)
    products: list[tuple[str, Any]] = []
    deliveries: list[str] = []
    for row_idx in range(first_data, first_data + 40):
        name = ws.cell(row=row_idx, column=name_col).value
        if name is None or str(name).strip() == "":
            break
        name_s = str(name).strip()
        if "доставк" in name_s.lower():
            deliveries.append(name_s)
            continue
        products.append((name_s, ws.cell(row=row_idx, column=sum_col).value))
    return products, deliveries


def _all_cell_texts(ws) -> list[str]:
    texts: list[str] = []
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.strip():
                texts.append(cell.value)
    return texts


def _vat_formula(ws) -> str:
    for row in ws.iter_rows(min_row=1, max_row=80, max_col=8):
        for cell in row:
            if cell.value == "в том числе НДС (22%)":
                _, _, sum_col, _ = _table_layout(ws)
                formula = ws.cell(row=cell.row, column=sum_col).value
                assert isinstance(formula, str), formula
                return formula
    raise AssertionError("VAT row not found")


def test_generate_xlsx_accepts_embed_flag_default_false() -> None:
    params = inspect.signature(generate_commercial_offer_xlsx).parameters
    assert "embed_delivery_in_unit_price" in params
    assert params["embed_delivery_in_unit_price"].default is False


def test_embed_xlsx_hides_delivery_and_keeps_grand_total() -> None:
    order = _discount_plate_order()
    totals = calculate_total_cost(order, discount_percent=50, logistics_cost=100.0)
    ws = _load_kp_sheet(
        order,
        discount_percent=50,
        logistics_cost=100.0,
        embed_delivery_in_unit_price=True,
    )
    products, deliveries = _product_and_delivery_rows(ws)
    assert deliveries == []
    amounts = [float(value) for _, value in products]
    assert abs(sum(amounts) - totals["total_with_vat"]) < 0.01
    assert abs(sum(amounts) - 161.0) < 0.01


def test_embed_xlsx_vat_formula_covers_all_product_rows() -> None:
    order = _discount_plate_order()
    ws = _load_kp_sheet(
        order,
        discount_percent=50,
        logistics_cost=100.0,
        embed_delivery_in_unit_price=True,
    )
    products, _ = _product_and_delivery_rows(ws)
    formula = _vat_formula(ws)
    product_sum = sum(float(value) for _, value in products)
    assert "*0.22" in formula.replace(" ", "") or f"*{VAT_RATE}" in formula.replace(" ", "")
    assert "SUM(" in formula.upper()
    assert abs(product_sum * VAT_RATE - 161.0 * 0.22) < 0.01


def test_embed_xlsx_adds_delivery_footnote() -> None:
    order = _discount_plate_order()
    ws = _load_kp_sheet(
        order,
        discount_percent=50,
        logistics_cost=100.0,
        embed_delivery_in_unit_price=True,
    )
    blob = "\n".join(_all_cell_texts(ws))
    assert "В стоимость изделий включена доставка" in blob
    assert "Скидка на доставку не распространяется" in blob


def test_plain_xlsx_still_has_delivery_row() -> None:
    order = _discount_plate_order()
    ws = _load_kp_sheet(order, discount_percent=50, logistics_cost=100.0)
    _, deliveries = _product_and_delivery_rows(ws)
    assert deliveries, "ordinary XLSX must keep the delivery line"
    blob = "\n".join(_all_cell_texts(ws))
    assert "В стоимость изделий включена доставка" not in blob


def test_embed_xlsx_mixed_pools_drop_both_delivery_rows(tmp_path: Path) -> None:
    db_path = _seed_pile_catalog(tmp_path)
    order = _mixed_order()
    totals = calculate_total_cost(
        order,
        discount_percent=0,
        logistics_cost=1000.0,
        pile_logistics_cost=2000.0,
        pile_catalog_db_path=db_path,
    )
    assert totals["plate_delivery_total"] > 0
    assert totals["pile_delivery_total"] > 0
    ws = _load_kp_sheet(
        order,
        logistics_cost=1000.0,
        pile_logistics_cost=2000.0,
        pile_catalog_db_path=db_path,
        embed_delivery_in_unit_price=True,
    )
    products, deliveries = _product_and_delivery_rows(ws)
    assert deliveries == []
    amounts = [float(value) for _, value in products]
    assert abs(sum(amounts) - totals["total_with_vat"]) < 0.01
    blob = "\n".join(_all_cell_texts(ws))
    assert "В стоимость изделий включена доставка" in blob


def test_embed_xlsx_keeps_delivery_row_for_unembedded_pool(
    tmp_path: Path, monkeypatch
) -> None:
    """Plate delivery > 0 with zero plate qty: keep that row, still embed piles."""
    import core.commercial_offer_xlsx as xlsx_mod

    real_totals = xlsx_mod.calculate_total_cost

    def _totals_with_orphan_plate_delivery(*args: Any, **kwargs: Any):
        totals = dict(real_totals(*args, **kwargs))
        totals["plate_delivery_total"] = 80.0
        totals["plate_trips"] = 1
        return totals

    monkeypatch.setattr(xlsx_mod, "calculate_total_cost", _totals_with_orphan_plate_delivery)
    db_path = _seed_pile_catalog(tmp_path)
    order = [
        {
            "line_id": "p0",
            "product_type": "plates",
            "name": "ПБ 60-12-8п",
            "mark": "ПБ 60-12-8п",
            "length_m": 1.0,
            "width_m": 1.0,
            "qty": 0,
            "unit_price": 1000.0,
            "concrete_grade": "М500",
        },
        {
            "line_id": "s1",
            "product_type": "piles",
            "product_kind": "pile",
            "name": "С60.30",
            "mark": "С60.30",
            "concrete_grade": "B25",
            "qty": 14,
            "unit_price": 50.0,
        },
    ]
    ws = _load_kp_sheet(
        order,
        pile_logistics_cost=1500.0,
        pile_catalog_db_path=db_path,
        embed_delivery_in_unit_price=True,
    )
    _, deliveries = _product_and_delivery_rows(ws)
    assert any("доставк" in name.lower() for name in deliveries)
    assert "Доставка свай" not in deliveries
    blob = "\n".join(_all_cell_texts(ws))
    assert "В стоимость изделий включена доставка" in blob
