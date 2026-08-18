"""Pure parser for fuel-operator transaction .xls exports (no DB / no app.*)."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Literal

import xlrd

ServiceType = Literal["fuel", "wash", "other"]

_WASH_SERVICE = "мойка"
_HEADER_MARKER = "Дата трн."
_FOOTER_PREFIX = "Итоги"
_RECONCILE_EPS = 0.01


@dataclass(frozen=True, slots=True)
class ParsedTxRow:
    """One transaction row from an operator export."""

    ts: datetime
    card_number: str
    service: str
    service_type: ServiceType
    fuel_grade: str | None
    qty_liters: float | None
    amount: float
    unit: str
    brand: str
    city: str
    raw_address: str


@dataclass(frozen=True, slots=True)
class ParsedTxFile:
    """Parsed .xls file with footer reconcile metadata."""

    filename: str
    rows: tuple[ParsedTxRow, ...]
    sum_liters: float
    sum_amount: float
    footer_liters: float | None
    footer_amount: float | None
    warnings: tuple[str, ...]


def classify_service(service: str) -> ServiceType:
    """Map raw «Услуга» text to fuel / wash / other."""
    text = (service or "").strip()
    lowered = text.lower()
    if lowered == _WASH_SERVICE:
        return "wash"
    if lowered.startswith("аи-"):
        return "fuel"
    return "other"


def _as_str(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float):
        if value == int(value):
            return str(int(value))
        return str(value)
    return str(value).strip()


def _parse_float(value: Any) -> float | None:
    if value in ("", None):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(_as_str(value).replace(",", "."))
    except ValueError:
        return None


def norm_card(value: Any) -> str:
    """Card number → string without trailing «.0» or spaces."""
    return _as_str(value).removesuffix(".0").strip()


def _parse_tx_datetime(cell: Any, datemode: int) -> datetime | None:
    if cell.ctype == xlrd.XL_CELL_DATE:
        return xlrd.xldate.xldate_as_datetime(cell.value, datemode)
    text = _as_str(cell.value)
    if not text:
        return None
    try:
        return datetime.fromisoformat(text)
    except ValueError:
        return None


def _parse_book(book: xlrd.Book, filename: str) -> ParsedTxFile:
    sheet = book.sheet_by_index(0)
    header_row: int | None = None
    for r in range(min(5, sheet.nrows)):
        if _as_str(sheet.cell_value(r, 0)) == _HEADER_MARKER:
            header_row = r
            break

    warnings: list[str] = []
    if header_row is None:
        warnings.append(f"{filename}: не найдена строка шапки «{_HEADER_MARKER}»")
        return ParsedTxFile(
            filename=filename,
            rows=(),
            sum_liters=0.0,
            sum_amount=0.0,
            footer_liters=None,
            footer_amount=None,
            warnings=tuple(warnings),
        )

    headers = [_as_str(sheet.cell_value(header_row, c)) for c in range(sheet.ncols)]
    col = {name: i for i, name in enumerate(headers) if name}

    def value(r: int, name: str, fallback: int) -> Any:
        i = col.get(name, fallback)
        if i is None or i >= sheet.ncols:
            return None
        return sheet.cell_value(r, i)

    rows: list[ParsedTxRow] = []
    sum_liters = 0.0
    sum_amount = 0.0
    footer_liters: float | None = None
    footer_amount: float | None = None

    for r in range(header_row + 1, sheet.nrows):
        first = _as_str(sheet.cell_value(r, 0))
        if not first:
            continue
        if first.startswith(_FOOTER_PREFIX):
            footer_liters = _parse_float(value(r, "Кол-во", 3))
            footer_amount = _parse_float(value(r, "Сумма с налогом, всего", 7))
            continue

        dt_col = col.get(_HEADER_MARKER, 0)
        ts = _parse_tx_datetime(sheet.cell(r, dt_col), book.datemode)
        if ts is None:
            continue

        service = _as_str(value(r, "Услуга", 2))
        service_type = classify_service(service)
        qty = _parse_float(value(r, "Кол-во", 3))
        amount = _parse_float(value(r, "Сумма с налогом, всего", 7)) or 0.0
        sum_liters += qty or 0.0
        sum_amount += amount

        rows.append(
            ParsedTxRow(
                ts=ts,
                card_number=norm_card(value(r, "Карта", 1)),
                service=service,
                service_type=service_type,
                fuel_grade=service if service_type == "fuel" else None,
                qty_liters=qty,
                amount=amount,
                unit=_as_str(value(r, "Ед. изм.", 4)),
                brand=_as_str(value(r, "Бренд", 9)),
                city=_as_str(value(r, "Город", 10)),
                raw_address=_as_str(value(r, "Адрес ТО", 11)),
            )
        )

    if footer_liters is not None and abs(sum_liters - footer_liters) > _RECONCILE_EPS:
        warnings.append(
            f"{filename}: Кол-во {sum_liters:.2f} ≠ Итоги {footer_liters:.2f}"
        )
    if footer_amount is not None and abs(sum_amount - footer_amount) > _RECONCILE_EPS:
        warnings.append(
            f"{filename}: Сумма {sum_amount:.2f} ≠ Итоги {footer_amount:.2f}"
        )

    return ParsedTxFile(
        filename=filename,
        rows=tuple(rows),
        sum_liters=sum_liters,
        sum_amount=sum_amount,
        footer_liters=footer_liters,
        footer_amount=footer_amount,
        warnings=tuple(warnings),
    )


def parse_transactions_xls(path: Path | str) -> ParsedTxFile:
    """Parse one operator .xls; footer mismatch → warnings, not exception."""
    file_path = Path(path)
    book = xlrd.open_workbook(filename=str(file_path))
    return _parse_book(book, file_path.name)


def parse_transactions_content(content: bytes, *, filename: str) -> ParsedTxFile:
    """Parse .xls bytes (same rules as ``parse_transactions_xls``)."""
    book = xlrd.open_workbook(file_contents=content)
    return _parse_book(book, filename)


def parse_transactions_directory(tx_dir: Path | str) -> tuple[list[ParsedTxRow], list[str]]:
    """Load all ``*.xls`` under a directory (script-compatible helper)."""
    directory = Path(tx_dir)
    all_rows: list[ParsedTxRow] = []
    all_warnings: list[str] = []
    for path in sorted(directory.glob("*.xls")):
        parsed = parse_transactions_xls(path)
        all_rows.extend(parsed.rows)
        all_warnings.extend(parsed.warnings)
    return all_rows, all_warnings
