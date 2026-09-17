"""GPS-004: parser for 1C «Прайс-лист» .xls exports."""

from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook

from core.pricelist_1c_parser import (
    Pricelist1CFormatError,
    PricelistRow,
    _parse_pricelist_book,
    parse_pricelist_xls,
)

REPO_ROOT = Path(__file__).resolve().parents[1]
GUID_DIR = REPO_ROOT / "банк знаний" / "Новая папка" / "GUID"

PRICE_FILES = {
    "Прайс сваи.xls": 344,
    "Прайс сваи мостовые.xls": 66,
    "Прайс сваи мостовые составные.xls": 4,
    "Прайс сваи составные.xls": 72,
    "Прайс блоки.xls": 18,
    "Прайс ЛМ, ЛП, ЛС, Перемычки, Прогоны, Тротуарные плиты, Фундаменты.xls": 200,
    "Прайс плиты.xls": 59631,
}

KNOWN_PILE = PricelistRow(
    guid="282b2926-ee51-11e9-8f4e-04d9f5387845",
    name="Сваи С 30.30-3",
    price=None,
    unit="шт",
    source_file="Прайс сваи.xls",
    row_index=3,
)
KNOWN_PILE_PRICED = PricelistRow(
    guid="a74cdadd-9e9a-11eb-8f7e-04d9f5387845",
    name="Сваи С 40.20-3",
    price=4800.0,
    unit="шт",
    source_file="Прайс сваи.xls",
    row_index=159,
)

_HEADER = [
    [
        "№",
        "Уникальный идентификатор (Номенклатура)",
        "Уникальный идентификатор (Характеристика)",
        "Товар",
        "Продажная",
        "",
        "",
        "",
        "",
        "",
    ],
    [
        "",
        "",
        "",
        "",
        "Старая цена",
        "Изменение",
        "%",
        "Цена",
        "Ед. изм.",
        "Уникальный идентификатор (Единица измерения)",
    ],
]


class _FakeSheet:
    def __init__(self, rows: list[list[object]]) -> None:
        self._rows = rows
        self.nrows = len(rows)
        self.ncols = max((len(row) for row in rows), default=0)

    def cell_value(self, row: int, col: int) -> object:
        if row >= self.nrows or col >= len(self._rows[row]):
            return ""
        return self._rows[row][col]


class _FakeBook:
    def __init__(self, sheets: dict[str, _FakeSheet]) -> None:
        self._sheets = sheets

    def sheet_names(self) -> list[str]:
        return list(self._sheets)

    def sheet_by_name(self, name: str) -> _FakeSheet:
        return self._sheets[name]


def _require_guid_dir() -> Path:
    if not GUID_DIR.is_dir():
        pytest.skip(f"нет каталога {GUID_DIR}")
    return GUID_DIR


def _require_file(name: str) -> Path:
    path = _require_guid_dir() / name
    if not path.is_file():
        pytest.skip(f"нет файла {path}")
    return path


def test_parse_piles_known_guid_and_name() -> None:
    rows = parse_pricelist_xls(_require_file("Прайс сваи.xls"))
    assert rows[0] == KNOWN_PILE
    priced = next(row for row in rows if row.guid == KNOWN_PILE_PRICED.guid)
    assert priced == KNOWN_PILE_PRICED


@pytest.mark.parametrize("filename,expected", list(PRICE_FILES.items()))
def test_parse_all_real_pricelist_files(filename: str, expected: int) -> None:
    rows = parse_pricelist_xls(_require_file(filename))
    assert len(rows) == expected
    assert all(row.guid and row.name for row in rows)
    assert all(row.source_file == filename for row in rows)
    assert all(row.row_index >= 3 for row in rows)
    assert rows[0].unit


def test_rejects_xlsx_even_with_pricelist_sheet(tmp_path: Path) -> None:
    path = tmp_path / "fake.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Прайс-лист"
    ws.append(_HEADER[0])
    ws.append(_HEADER[1])
    ws.append(["1", "guid", "", "Товар", "", "", 0, 100, "шт", ""])
    wb.save(path)

    with pytest.raises(Pricelist1CFormatError, match="это не выгрузка Прайс-лист 1С"):
        parse_pricelist_xls(path)

    xls_named = tmp_path / "fake.xls"
    xls_named.write_bytes(path.read_bytes())
    with pytest.raises(Pricelist1CFormatError, match="это не выгрузка Прайс-лист 1С"):
        parse_pricelist_xls(xls_named)


def test_rejects_garbage_bytes(tmp_path: Path) -> None:
    path = tmp_path / "broken.xls"
    path.write_bytes(b"this is not an excel file")
    with pytest.raises(Pricelist1CFormatError, match="это не выгрузка Прайс-лист 1С"):
        parse_pricelist_xls(path)


def test_missing_file_is_file_not_found(tmp_path: Path) -> None:
    with pytest.raises(FileNotFoundError):
        parse_pricelist_xls(tmp_path / "нет такого.xls")


def test_wrong_sheet_human_readable() -> None:
    book = _FakeBook({"Лист1": _FakeSheet(_HEADER)})
    with pytest.raises(Pricelist1CFormatError, match="это не выгрузка Прайс-лист 1С"):
        _parse_pricelist_book(book, "other.xls")


def test_wrong_headers_human_readable() -> None:
    rows = [
        ["A", "B", "C", "D", "E", "F", "G", "H", "I"],
        ["1", "2", "3", "4", "5", "6", "7", "8", "9"],
        ["1", "guid-1", "", "Name", "", "", 0, 10, "шт"],
    ]
    book = _FakeBook({"Прайс-лист": _FakeSheet(rows)})
    with pytest.raises(Pricelist1CFormatError, match="это не выгрузка Прайс-лист 1С"):
        _parse_pricelist_book(book, "other.xls")


def test_skips_empty_guid_name_and_trailing_rows() -> None:
    rows = [
        *_HEADER,
        ["1", "AAA-BBB", "", "Сваи С 30.30-3", "", "", 0, "", "шт", ""],
        ["", "", "", "", "", "", "", "", "", ""],
        ["2", "", "", "Без GUID", "", "", 0, 1, "шт", ""],
        ["3", "CCC-DDD", "", "", "", "", 0, 2, "шт", ""],
        ["4", "EEE-FFF", "", "Сваи С 50.30-6", "", "", 0, "6700", "шт", ""],
    ]
    parsed = _parse_pricelist_book(
        _FakeBook({"Прайс-лист": _FakeSheet(rows)}),
        "mini.xls",
    )
    assert [row.name for row in parsed] == ["Сваи С 30.30-3", "Сваи С 50.30-6"]
    assert parsed[0].guid == "aaa-bbb"
    assert parsed[0].price is None
    assert parsed[1].price == 6700.0
    assert parsed[1].row_index == 7
