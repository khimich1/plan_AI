"""Preview/apply factory price Excel into pb.db (one parser per filename)."""

from __future__ import annotations

import hashlib
import sqlite3
import tempfile
from pathlib import Path
from typing import Callable, Sequence, cast

from app.core.settings import get_settings
from app.schemas.price_desk import (
    STATUS_KINDS,
    PriceDeskDiffExample,
    PriceDeskDiffExamples,
    PriceDeskKind,
    PriceDeskPreview,
    PriceDeskStatus,
    PriceGroupStatus,
)
from core.bridge_pile_price_db import (
    import_bridge_pile_prices_from_xlsx,
    init_bridge_pile_prices_schema,
    parse_bridge_pile_price_rows_from_xlsx,
)
from core.fbs_price_db import (
    import_fbs_prices_from_xlsx,
    init_fbs_prices_schema,
    parse_fbs_price_rows_from_xlsx,
)
from core.march_price_db import (
    import_march_prices_from_xlsx,
    init_march_prices_schema,
    parse_march_price_rows_from_xlsx,
)
from core.pile_price_db import (
    import_pile_prices_from_xlsx,
    init_pile_prices_schema,
    parse_pile_price_rows_from_xlsx,
)
from core.plate_nikita_parser import find_price_sheet_name, has_nikita_column, parse_plate_nikita_rows
from core.price_db import import_plate_nikita_from_xlsx, init_schema, list_plate_prices, plate_prices_meta
from core.price_desk_classify import (
    MSG_UNKNOWN_KIND,
    ClassifyError,
    classify_price_filename,
    parse_price_list_date_from_filename,
)
from core.price_desk_diff import diff_price_maps, prices_to_map
from core.step_price_db import (
    import_step_prices_from_xlsx,
    init_step_prices_schema,
    parse_step_price_rows_from_xlsx,
)

MSG_COMPOSITE = "группа пока не ведётся"
MSG_NO_SHEET = "лист «Прайс» не найден"
MSG_NO_NIKITA = "это не прайс ПБ: нет колонки Никиты"
MSG_SHA_MISMATCH = "файл не тот, что в превью"
MSG_NEED_SHA = "укажите file_sha256 с превью"
MSG_EMPTY_ROWS = "не удалось прочитать цены с листа «Прайс»"
MSG_NOT_PILE = "имя файла и содержимое не совпали: это не прайс цельных свай"
MSG_EMPTY_FILE = "Пустой файл."
MSG_NOT_EXCEL = "Нужен файл Excel (.xls или .xlsx)"

_KEY_LEN: dict[PriceDeskKind, int] = {
    "plates": 2,
    "fbs": 2,
    "march": 2,
    "bridge_pile": 2,
    "pile": 2,
    "step": 1,
}

_META_SQL: dict[PriceDeskKind, tuple[Callable[[str], None], str]] = {
    "plates": (init_schema, "prices"),
    "fbs": (init_fbs_prices_schema, "fbs_prices"),
    "march": (init_march_prices_schema, "march_prices"),
    "step": (init_step_prices_schema, "step_prices"),
    "bridge_pile": (init_bridge_pile_prices_schema, "bridge_pile_prices"),
    "pile": (init_pile_prices_schema, "pile_prices"),
}


class PriceDeskError(Exception):
    """Client-facing desk error (422/409)."""

    def __init__(self, message: str, status_code: int = 422) -> None:
        super().__init__(message)
        self.message = message
        self.status_code = status_code


def file_sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _safe_upload_name(filename: str | None) -> str:
    raw = (filename or "").strip()
    base = Path(raw).name if raw else "upload.xlsx"
    if not base or base in {".", ".."}:
        return "upload.xlsx"
    return base


def _assert_excel_name(name: str) -> None:
    lower = name.lower()
    if not (lower.endswith(".xlsx") or lower.endswith(".xls")):
        raise PriceDeskError(MSG_NOT_EXCEL)


class PriceDeskService:
    def __init__(self, db_path: Path | str | None = None) -> None:
        self._db_path = Path(db_path) if db_path is not None else None

    def db_path(self) -> str:
        if self._db_path is not None:
            return str(self._db_path)
        return str(get_settings().pb_db_path)

    def status(self) -> PriceDeskStatus:
        db = self.db_path()
        groups: list[PriceGroupStatus] = []
        for kind in STATUS_KINDS:
            row_count, list_date, imported_at = self._group_meta(kind, db)
            groups.append(
                PriceGroupStatus(
                    product_kind=kind,
                    price_list_date=list_date,
                    imported_at=imported_at,
                    row_count=row_count,
                )
            )
        return PriceDeskStatus(groups=groups)

    def preview(self, file_bytes: bytes, filename: str | None) -> PriceDeskPreview:
        return self._run(file_bytes, filename, apply=False, expected_sha256=None)

    def apply(
        self,
        file_bytes: bytes,
        filename: str | None,
        file_sha256_hex: str | None,
    ) -> PriceDeskPreview:
        expected = (file_sha256_hex or "").strip().lower()
        if not expected:
            raise PriceDeskError(MSG_NEED_SHA)
        return self._run(file_bytes, filename, apply=True, expected_sha256=expected)

    def _run(
        self,
        file_bytes: bytes,
        filename: str | None,
        *,
        apply: bool,
        expected_sha256: str | None,
    ) -> PriceDeskPreview:
        if not file_bytes:
            raise PriceDeskError(MSG_EMPTY_FILE)
        name = _safe_upload_name(filename)
        _assert_excel_name(name)
        digest = file_sha256(file_bytes)
        if expected_sha256 is not None and digest != expected_sha256:
            raise PriceDeskError(MSG_SHA_MISMATCH, status_code=409)

        try:
            kind = classify_price_filename(name)
        except ClassifyError as exc:
            raise PriceDeskError(str(exc) or MSG_UNKNOWN_KIND) from exc
        if kind == "composite":
            raise PriceDeskError(MSG_COMPOSITE)
        product_kind = cast(PriceDeskKind, kind)

        db = self.db_path()
        with tempfile.TemporaryDirectory(prefix="price-desk-") as tmp:
            path = Path(tmp) / name
            path.write_bytes(file_bytes)
            parsed = self._parse_and_validate(product_kind, str(path))
            preview = self._preview_from_rows(product_kind, parsed, digest, name)
            if apply:
                self._import_kind(product_kind, str(path), db)
            return preview

    def _parse_and_validate(self, kind: PriceDeskKind, path: str) -> Sequence[tuple]:
        if find_price_sheet_name(path) is None:
            raise PriceDeskError(MSG_NO_SHEET)
        if kind == "plates":
            if not has_nikita_column(path):
                raise PriceDeskError(MSG_NO_NIKITA)
            rows = parse_plate_nikita_rows(path)
        elif kind == "fbs":
            rows = parse_fbs_price_rows_from_xlsx(path)
        elif kind == "march":
            rows = parse_march_price_rows_from_xlsx(path)
        elif kind == "step":
            rows = parse_step_price_rows_from_xlsx(path)
        elif kind == "bridge_pile":
            rows = parse_bridge_pile_price_rows_from_xlsx(path)
        elif kind == "pile":
            rows = parse_pile_price_rows_from_xlsx(path)
            if any("ФБС" in str(row[0]).upper() for row in rows):
                raise PriceDeskError(MSG_NOT_PILE)
        else:
            raise PriceDeskError(MSG_UNKNOWN_KIND)
        if not rows:
            raise PriceDeskError(MSG_EMPTY_ROWS)
        return rows

    def _preview_from_rows(
        self,
        kind: PriceDeskKind,
        rows: Sequence[tuple],
        digest: str,
        filename: str,
    ) -> PriceDeskPreview:
        key_len = _KEY_LEN[kind]
        incoming = prices_to_map(rows, key_len=key_len)
        current_rows = self._current_rows(kind, self.db_path())
        current = prices_to_map(current_rows, key_len=key_len)
        diff = diff_price_maps(current, incoming)
        examples = PriceDeskDiffExamples(
            changed=[_to_schema_example(item) for item in diff.examples.changed],
            new=[_to_schema_example(item) for item in diff.examples.new],
            missing=[_to_schema_example(item) for item in diff.examples.missing],
            unchanged=[_to_schema_example(item) for item in diff.examples.unchanged],
        )
        return PriceDeskPreview(
            product_kind=kind,
            price_list_date=parse_price_list_date_from_filename(filename),
            file_sha256=digest,
            parsed_rows=len(rows),
            changed=diff.changed,
            new=diff.new,
            missing=diff.missing,
            unchanged=diff.unchanged,
            examples=examples,
        )

    def _import_kind(self, kind: PriceDeskKind, path: str, db: str) -> int:
        if kind == "plates":
            return import_plate_nikita_from_xlsx(path, db)
        if kind == "fbs":
            return import_fbs_prices_from_xlsx(path, db)
        if kind == "march":
            return import_march_prices_from_xlsx(path, db)
        if kind == "step":
            return import_step_prices_from_xlsx(path, db)
        if kind == "bridge_pile":
            return import_bridge_pile_prices_from_xlsx(path, db)
        if kind == "pile":
            return import_pile_prices_from_xlsx(path, db)
        raise PriceDeskError(MSG_UNKNOWN_KIND)

    def _current_rows(self, kind: PriceDeskKind, db: str) -> list[tuple]:
        if kind == "plates":
            return list_plate_prices(db)
        init_fn, table = _META_SQL[kind]
        init_fn(db)
        conn = sqlite3.connect(db)
        try:
            cur = conn.cursor()
            if kind == "step":
                cur.execute(f"SELECT mark, price FROM {table}")
                return [(str(r[0]), float(r[1])) for r in cur.fetchall()]
            cur.execute(f"SELECT mark, concrete_grade, price FROM {table}")
            return [(str(r[0]), str(r[1]), float(r[2])) for r in cur.fetchall()]
        finally:
            conn.close()

    def _group_meta(
        self, kind: PriceDeskKind, db: str
    ) -> tuple[int, str | None, str | None]:
        if kind == "plates":
            return plate_prices_meta(db)
        init_fn, table = _META_SQL[kind]
        init_fn(db)
        conn = sqlite3.connect(db)
        try:
            cur = conn.cursor()
            cur.execute(
                f"SELECT COUNT(*), MAX(price_list_date), MAX(imported_at) FROM {table}"
            )
            row = cur.fetchone() or (0, None, None)
            return int(row[0] or 0), row[1], row[2]
        finally:
            conn.close()


def _to_schema_example(item) -> PriceDeskDiffExample:
    return PriceDeskDiffExample(
        key=item.key,
        old_price=item.old_price,
        new_price=item.new_price,
    )
