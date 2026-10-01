"""Книга спецификации: бланк № 650, строки товара, подвал по якорям."""

from __future__ import annotations

import html
import io
import zipfile
from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

from core.commercial_pricing import vat_included_from_line_sums
from core.specification_text import (
    SpecificationChoice,
    SpecificationParagraphs,
    format_rub,
    render_specification,
    total_amount_words,
)
from core.specification_xlsx import (
    SpecificationHeader,
    SpecificationPayableLine,
    build_specification_xlsx,
)
from core.supply_contract import build_preamble

_BLANK = Path(__file__).resolve().parents[1] / "templates" / "specification.xlsx"
_FORBIDDEN = ("ЭТАЛОН", "9726099814", "3554802")
_SPEC_DATE = date(2026, 9, 18)
_CONTRACT_DATE = date(2026, 9, 1)
_LINES = (
    SpecificationPayableLine(name="ПБ 60-12-8п", qty=2, price=Decimal("100.50")),
    SpecificationPayableLine(name="ПБ 45-12-8п", qty=1, price=Decimal("50.00")),
)


def _preamble() -> str:
    return build_preamble(
        legal_form="ooo",
        full_name="Ромашка",
        signatory_name="Иванов Иван Иванович",
        signatory_position="директора",
        authority_basis="устав",
    )


def _header(*, invoice_number: str | None, concrete_grade: str | None = None) -> SpecificationHeader:
    return SpecificationHeader(
        invoice_number=invoice_number,
        spec_date=_SPEC_DATE,
        contract_number="0001/09/26",
        contract_date=_CONTRACT_DATE,
        buyer_preamble=_preamble(),
        buyer_short_name="Ромашка",
        buyer_inn="7604368921",
        buyer_kpp="760401001",
        buyer_signatory_position="Директор",
        buyer_signatory_short="Иванов И. И.",
        concrete_grade=concrete_grade,
    )


def _plate_paragraphs() -> SpecificationParagraphs:
    return render_specification(
        SpecificationChoice(
            payment="prepay_100",
            term="by_date",
            term_date=date(2026, 10, 1),
            delivery="pickup",
        ),
        payable_total=25_100,
        has_piles=False,
    )


def _share_paragraphs() -> SpecificationParagraphs:
    return render_specification(
        SpecificationChoice(
            payment="split_share",
            payment_date=_SPEC_DATE,
            second_payment_date=date(2026, 10, 1),
            second_share_percent=20,
            payment_days=5,
            term="by_date",
            term_date=date(2026, 10, 1),
            delivery="site",
            delivery_address="г. Кострома, ул. Лесная, 4",
        ),
        payable_total=25_100,
        has_piles=False,
    )


def _build(
    invoice_number: str | None,
    *,
    concrete_grade: str | None = None,
    paragraphs: SpecificationParagraphs | None = None,
    lines: list[SpecificationPayableLine] | None = None,
) -> bytes:
    return build_specification_xlsx(
        _header(invoice_number=invoice_number, concrete_grade=concrete_grade),
        list(_LINES if lines is None else lines),
        paragraphs or _plate_paragraphs(),
    )


def _xlsx_text(data: bytes) -> str:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        parts = [
            html.unescape(archive.read(name).decode("utf-8"))
            for name in archive.namelist()
            if name.endswith(".xml")
        ]
    return "\n".join(parts)


def _sheet(data: bytes):
    return load_workbook(io.BytesIO(data))["Лист_1"]


def _single_row_spans(sheet, row: int) -> list[tuple[int, int]]:
    return sorted(
        (merged.min_col, merged.max_col)
        for merged in sheet.merged_cells.ranges
        if merged.min_row == row and merged.max_row == row
    )


def _item_rows(sheet) -> list[int]:
    return [
        row
        for row in range(1, sheet.max_row + 1)
        if len(_single_row_spans(sheet, row)) >= 6
    ]


def _exact(sheet, text: str):
    for row in sheet.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            if isinstance(cell.value, str) and cell.value.strip() == text:
                return cell
    raise AssertionError(text)


def _containing(sheet, needle: str):
    found = None
    for row in sheet.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell) or not isinstance(cell.value, str):
                continue
            if needle in cell.value and (found is None or len(cell.value) < len(str(found.value))):
                found = cell
    if found is None:
        raise AssertionError(needle)
    return found


def _amount_near(sheet, label: str) -> float:
    anchor = _containing(sheet, label)
    for cell in sheet[anchor.row]:
        if isinstance(cell.value, (int, float)) and cell.column > anchor.column:
            return float(cell.value)
    raise AssertionError(label)


def _prototype_spans() -> list[tuple[int, int]]:
    sheet = load_workbook(_BLANK)["Лист_1"]
    rows = _item_rows(sheet)
    assert len(rows) == 1
    return _single_row_spans(sheet, rows[0])


def test_blank_has_one_item_row_footer_anchors_and_no_deal_650() -> None:
    text = _xlsx_text(_BLANK.read_bytes())
    for token in _FORBIDDEN:
        assert token not in text

    sheet = load_workbook(_BLANK)["Лист_1"]
    assert sheet.page_setup.fitToWidth == 1
    assert sheet.sheet_properties.pageSetUpPr.fitToPage is True
    assert len(_item_rows(sheet)) == 1
    assert sheet.column_dimensions["B"].width == pytest.approx(3.5, abs=0.05)
    assert sheet.max_column >= 48
    for anchor in (
        "Итого:",
        "В т.ч. НДС (22%):",
        "Итого с НДС:",
        "Всего наименований",
        "Условия оплаты:",
        "Условия поставки:",
        "Вид поставки:",
        "Поставщик",
        "Покупатель",
        "Шишов А. В.",
        "4401082520",
        "г. Кострома",
    ):
        assert anchor in text
    supplier = _exact(sheet, "Поставщик")
    buyer = _exact(sheet, "Покупатель")
    assert supplier.row == buyer.row
    assert supplier.column != buyer.column


def test_supplier_is_zhbk_start_not_the_offer_letterhead() -> None:
    text = _xlsx_text(_build(None))

    assert "4401082520" in text
    assert "440101001" in text
    assert "ООО «ЖБК СТАРТ»" in text
    assert "Шишов А. В." in text
    assert "4705123456" not in text
    assert "Комбинат ЖБК" not in text
    assert "760345001" not in text


def test_empty_and_filled_invoice_number_keep_the_passed_date() -> None:
    empty = _xlsx_text(_build(None))
    numbered = _xlsx_text(_build("ЯР-15"))

    assert "Спецификация по счету № _____" in empty
    assert "от 18.09.2026г." in empty
    assert "к Договору поставки № 0001/09/26 от 01.09.2026г." in empty
    assert "от года" not in empty
    assert "№ ЯР-15" in numbered
    assert "_____" not in numbered
    assert "от 18.09.2026г." in numbered
    assert "г. Кострома" in empty
    assert "18 сентября 2026\u00a0г." in empty
    assert "18 сентября 2026\u00a0г." in numbered


def test_buyer_preamble_is_not_wrapped_in_a_second_v_litse() -> None:
    preamble = _preamble()
    text = _xlsx_text(_build(None))

    assert preamble.count("в лице") == 1
    assert preamble in text
    assert text.count("в лице") == preamble.count("в лице") + 1
    assert "в лице в лице" not in text
    assert "настоящую Спецификацию" in text
    assert "изготовить и отгрузить" in text
    assert "ООО Ромашка" in text
    assert "ООО ООО" not in text
    assert "Иванов И. И." in text
    assert text.count("М.П.") >= 2


def test_totals_match_line_sums_and_included_vat() -> None:
    data = _build(None)
    sheet = _sheet(data)
    line_sums = [Decimal("201.00"), Decimal("50.00")]
    vat = vat_included_from_line_sums(line_sums)
    total = sum(line_sums, Decimal("0.00"))

    assert _amount_near(sheet, "Итого:") == float(total)
    assert _amount_near(sheet, "Итого с НДС:") == float(total)
    assert _amount_near(sheet, "В т.ч. НДС (22%):") == float(vat)

    names = [row for row in sheet.iter_rows() if any(cell.value == "ПБ 60-12-8п" for cell in row)]
    assert len(names) == 1
    values = [cell.value for cell in names[0] if cell.value is not None]
    assert 2 in values
    assert "шт" in values
    assert 22 in values
    assert pytest.approx(201.0) in values
    total_kop = int(total * 100)
    sentence = _containing(sheet, "Всего наименований").value
    assert sentence == (
        f"Всего наименований 2, на сумму {format_rub(total_kop)} "
        f"({total_amount_words(total_kop)}), в т.ч. НДС — {format_rub(int(vat * 100))} руб."
    )


def test_condition_cells_keep_the_scheme_paragraph_under_the_blank_label() -> None:
    paragraphs = _share_paragraphs()
    data = _build(None, paragraphs=paragraphs)
    sheet = _sheet(data)
    text = _xlsx_text(data)

    payment = _containing(sheet, "Условия оплаты:")
    delivery = _containing(sheet, "Условия поставки:")
    kind = _containing(sheet, "Вид поставки:")
    assert str(payment.value).startswith("Условия оплаты: ")
    assert paragraphs.payment in str(payment.value)
    assert str(delivery.value).startswith("Условия поставки: ")
    assert paragraphs.term in str(delivery.value)
    assert str(kind.value).startswith("Вид поставки: ")
    assert paragraphs.delivery in str(kind.value)
    assert "50%" in str(payment.value)
    assert "2000000" not in text
    assert "1554802" not in text


def test_plate_sheet_has_no_concrete_grade_and_no_pile_rhythm() -> None:
    text = _xlsx_text(_build(None))

    assert "Марка бетона" not in text
    assert "свай в неделю" not in text
    assert "свай в день" not in text

    graded = _xlsx_text(_build("ЯР-15", concrete_grade="В25"))
    assert "Марка бетона" not in graded
    assert "В25" not in graded


def test_one_line_keeps_the_prototype_merges_and_puts_parties_on_one_row() -> None:
    lines = [SpecificationPayableLine(name="ПБ 16-12-8п", qty=1, price=Decimal("10.00"))]
    data = _build(None, lines=lines)
    sheet = _sheet(data)
    text = _xlsx_text(data)
    rows = _item_rows(sheet)
    pattern = _prototype_spans()

    assert sheet.title == "Лист_1"
    assert sheet.page_setup.fitToWidth == 1
    assert rows == [rows[0]]
    assert _single_row_spans(sheet, rows[0]) == pattern
    assert _containing(sheet, "Условия оплаты:").row > rows[-1]
    supplier = _exact(sheet, "Поставщик")
    buyer = _exact(sheet, "Покупатель")
    assert supplier.row == buyer.row
    assert supplier.column != buyer.column
    for token in _FORBIDDEN:
        assert token not in text


def test_forty_lines_shift_the_prototype_merges_and_the_footer() -> None:
    lines = [
        SpecificationPayableLine(name=f"Позиция {index}", qty=1, price=Decimal("10.00"))
        for index in range(1, 41)
    ]
    sheet = _sheet(_build(None, lines=lines))
    rows = _item_rows(sheet)
    pattern = _prototype_spans()

    assert len(rows) == 40
    assert rows[-1] - rows[0] == 39
    assert _single_row_spans(sheet, rows[0]) == pattern
    assert _single_row_spans(sheet, rows[-1]) == pattern
    assert _containing(sheet, "Условия оплаты:").row > rows[-1]
    assert _exact(sheet, "Поставщик").row == _exact(sheet, "Покупатель").row
    assert _containing(sheet, "Итого:").row > rows[-1]
    assert "Позиция 1" in [cell.value for row in sheet.iter_rows() for cell in row]
    assert "Позиция 40" in [cell.value for row in sheet.iter_rows() for cell in row]
