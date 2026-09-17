#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Парсер выгрузки 1С «Прайс-лист» (.xls BIFF).

Пустая GUID или пустое наименование: строка пропускается (хвост файла,
неполный ряд). В результат не попадает. Дубли имён/GUID не разрешаются —
парсер отдаёт все валидные строки как есть.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

import xlrd
from xlrd import XLRDError

SHEET_NAME = "Прайс-лист"

# 0-based; в Excel: GUID=кол.2, товар=4, цена=8, ед.=9. Данные с 3-й строки.
_COL_GUID = 1
_COL_NAME = 3
_COL_PRICE = 7
_COL_UNIT = 8
_HEADER_ROWS = 2
_MIN_COLS = 9

_FORMAT_PREFIX = "это не выгрузка Прайс-лист 1С"


class Pricelist1CFormatError(ValueError):
    """Файл не является выгрузкой 1С «Прайс-лист»."""


@dataclass(frozen=True, slots=True)
class PricelistRow:
    guid: str
    name: str
    price: float | None
    unit: str
    source_file: str
    row_index: int


def parse_pricelist_xls(path: Path | str) -> list[PricelistRow]:
    """Прочитать .xls выгрузки 1С «Прайс-лист».

    Raises:
        FileNotFoundError: путь не существует.
        Pricelist1CFormatError: не BIFF .xls или заголовок не похож на выгрузку 1С.
    """
    file_path = Path(path)
    if not file_path.is_file():
        raise FileNotFoundError(str(file_path))
    try:
        book = xlrd.open_workbook(filename=str(file_path), formatting_info=False)
    except XLRDError as exc:
        raise Pricelist1CFormatError(_open_error_message(exc)) from exc
    except Exception as exc:
        if isinstance(exc, (FileNotFoundError, Pricelist1CFormatError)):
            raise
        raise Pricelist1CFormatError(
            f"{_FORMAT_PREFIX}: не удалось прочитать файл ({exc})"
        ) from exc
    return _parse_pricelist_book(book, file_path.name)


def _parse_pricelist_book(book: Any, source_file: str) -> list[PricelistRow]:
    names = list(book.sheet_names())
    if SHEET_NAME not in names:
        raise Pricelist1CFormatError(
            f"{_FORMAT_PREFIX}: нет листа «{SHEET_NAME}»"
        )
    sheet = book.sheet_by_name(SHEET_NAME)
    _require_pricelist_header(sheet)

    rows: list[PricelistRow] = []
    for r in range(_HEADER_ROWS, sheet.nrows):
        guid = _as_str(sheet.cell_value(r, _COL_GUID)).lower()
        name = _as_str(sheet.cell_value(r, _COL_NAME))
        if not guid or not name:
            continue
        unit = _as_str(sheet.cell_value(r, _COL_UNIT))
        rows.append(
            PricelistRow(
                guid=guid,
                name=name,
                price=_parse_price(sheet.cell_value(r, _COL_PRICE)),
                unit=unit,
                source_file=source_file,
                row_index=r + 1,
            )
        )
    return rows


def _require_pricelist_header(sheet: Any) -> None:
    if sheet.nrows < _HEADER_ROWS or sheet.ncols < _MIN_COLS:
        raise Pricelist1CFormatError(
            f"{_FORMAT_PREFIX}: слишком мало строк или колонок"
        )
    guid_h = _norm_header(sheet.cell_value(0, _COL_GUID))
    name_h = _norm_header(sheet.cell_value(0, _COL_NAME))
    price_group = _norm_header(sheet.cell_value(0, 4))
    price_h = _norm_header(sheet.cell_value(1, _COL_PRICE))
    unit_h = _norm_header(sheet.cell_value(1, _COL_UNIT))
    if "уникальный идентификатор" not in guid_h or "номенклатура" not in guid_h:
        raise Pricelist1CFormatError(
            f"{_FORMAT_PREFIX}: нет колонки GUID номенклатуры"
        )
    if name_h != "товар":
        raise Pricelist1CFormatError(f"{_FORMAT_PREFIX}: нет колонки «Товар»")
    if "продажная" not in price_group:
        raise Pricelist1CFormatError(f"{_FORMAT_PREFIX}: нет вида цены «Продажная»")
    if price_h != "цена":
        raise Pricelist1CFormatError(
            f"{_FORMAT_PREFIX}: нет колонки цены «Продажная»"
        )
    if not unit_h.startswith("ед"):
        raise Pricelist1CFormatError(f"{_FORMAT_PREFIX}: нет колонки единицы измерения")


def _open_error_message(exc: BaseException) -> str:
    text = str(exc).lower()
    if "xlsx" in text:
        return f"{_FORMAT_PREFIX}: нужен файл .xls (выгрузка 1С), не .xlsx"
    return f"{_FORMAT_PREFIX}: не удалось прочитать файл ({exc})"


def _as_str(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float):
        if value == int(value):
            return str(int(value))
        return str(value).strip()
    return str(value).strip()


def _norm_header(value: Any) -> str:
    text = _as_str(value).replace("ё", "е").replace("Ё", "Е").lower()
    return " ".join(text.split())


def _parse_price(value: Any) -> float | None:
    if value in ("", None):
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        number = float(value)
        return number if number > 0 else None
    text = _as_str(value).replace(" ", "").replace(",", ".")
    if not text:
        return None
    try:
        number = float(text)
    except ValueError:
        return None
    return number if number > 0 else None
