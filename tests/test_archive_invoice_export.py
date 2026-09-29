"""Сборщик «Счёт на оплату» и первая выгрузка в папку обмена."""

from __future__ import annotations

import json
import sqlite3
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import pytest

from core.guid_gate import (
    GateReport,
    InvoiceGuidBlockError,
    MissingItem,
    OrderLine,
    ReadyItem,
)
from core.supply_contract import attach_contract_number


def _offer(**overrides: object) -> dict:
    base: dict = {
        "kp_id": 22,
        "date": "2026-09-11",
        "manager": "Пургина Ольга Владимировна",
        "discount_percent": 13.0,
        "discount_amount": 111.11,
        "vat_amount": 1.23,
        "total_amount": 999.99,
        "delivery": 18600.0,
        "counterparty": {
            "name": "АКТИВСТРОЙ ООО",
            "code": "00-00000884",
            "guid": None,
            "inn": "7604368921",
            "kpp": "760401001",
        },
        "lines": [
            {
                "name": "Плиты ПБ 45,4-12-8п",
                "qty": 11,
                "price": 15421.0,
                "discount_percent": 13.0,
                "discount_amount": 50.5,
                "sum": 40.0,
                "vat_amount": 7.77,
                "sum_with_vat": 47.77,
                "guid": "791546a2-0a50-11ec-8f8f-04d9f5387845",
            }
        ],
    }
    base.update(overrides)
    return base


_WAREHOUSE = "Склад Готовой Продукции"


def _active_contract() -> dict:
    return {
        "counterparty_id": 7,
        "status": "нет",
        "number": "0001/09/26",
    }


def test_invoice_document_name_event_and_contract_stamp() -> None:
    from core.invoice_export import build_invoice_document

    document, _digest = build_invoice_document(_offer(), action="create")
    stamped = attach_contract_number(document, 7, _active_contract())

    assert stamped["Документ"] == "Счёт на оплату"
    assert stamped["event"] == "invoice_export"
    assert stamped["action"] == "create"
    keys = list(stamped)
    assert keys.index("НомерДоговора") == keys.index("Контрагент") + 1
    assert stamped["НомерДоговора"] == "0001/09/26"
    assert "НомерДоговора" not in document


def test_invoice_create_omits_order_number_and_update_requires_it() -> None:
    from core.invoice_export import InvoiceBuildError, build_invoice_document

    created, _digest = build_invoice_document(
        _offer(),
        action="create",
        order_number="ЯР-0001467",
    )
    assert "НомерЗаказа" not in created

    with pytest.raises(InvoiceBuildError):
        build_invoice_document(_offer(), action="update", order_number=None)
    with pytest.raises(InvoiceBuildError):
        build_invoice_document(_offer(), action="update", order_number="  ")

    updated, _updated_digest = build_invoice_document(
        _offer(),
        action="update",
        order_number="ЯР-0001467",
    )
    assert updated["НомерЗаказа"] == "ЯР-0001467"
    assert updated["action"] == "update"


def _payable(price: str, discount_percent: str, qty: int, delivery_rub: str = "0") -> Decimal:
    product = Decimal(price) * (Decimal(1) - Decimal(discount_percent) / Decimal(100)) * qty
    return (product + Decimal(delivery_rub)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def _included_vat(payable: Decimal) -> Decimal:
    rate = Decimal(22) / Decimal(122)
    return (payable * rate).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def test_invoice_header_totals_follow_lines_not_card_vat() -> None:
    from core.invoice_export import build_invoice_document

    document, _digest = build_invoice_document(_offer(), action="create")
    payable = _payable("15421", "13", 11)

    assert document["ПроцентСкидки"] == 13.0
    assert document["СуммаСкидки"] == 111.11
    assert document["Самовывоз"] is True
    assert document["СуммаДокумента"] == float(payable)
    assert document["СуммаНДС"] == float(_included_vat(payable))
    assert document["СуммаНДС"] != 1.23
    assert document["СуммаДокумента"] != 999.99
    line = document["Товары"][0]
    assert line["Количество"] == 11
    assert line["Цена"] == 15421.0
    assert line["ПроцентСкидки"] == 13.0
    assert line["СуммаДоставки"] == 0.0
    assert line["СуммаСкидки"] == 50.5
    assert line["Сумма"] == 40.0
    assert line["СуммаНДС"] == 7.77
    assert line["СуммаСНДС"] == 47.77
    assert "Склад" not in line
    assert line["Номенклатура"] == "Плиты ПБ 45,4-12-8п"
    assert document["НомерВПриложении"] == 22
    assert document["Контрагент"]["ИНН"] == "7604368921"
    assert document["Контрагент"]["Код"] == "00-00000884"
    assert document["Менеджер"] == "Пургина Ольга Владимировна"


def test_invoice_pickup_keeps_pre_discount_price() -> None:
    from core.invoice_export import build_invoice_document

    document, _digest = build_invoice_document(
        _offer(
            discount_percent=10.0,
            lines=[
                {
                    "name": "Плиты ПБ 65-12-8п",
                    "qty": 2,
                    "price": 23643.0,
                    "discount_percent": 10.0,
                    "guid": "plate-guid",
                }
            ],
        ),
        action="create",
        warehouse="Склад Готовой Продукции",
        line_delivery_kopecks=[0],
    )

    line = document["Товары"][0]
    assert document["Самовывоз"] is True
    assert document["Склад"] == "Склад Готовой Продукции"
    assert line["Цена"] == 23643.0
    assert line["ПроцентСкидки"] == 10.0
    assert line["СуммаДоставки"] == 0.0
    assert "Склад" not in line
    assert document["СуммаДокумента"] == float(_payable("23643", "10", 2))


def test_invoice_delivery_share_is_not_discounted() -> None:
    from core.invoice_export import build_invoice_document

    line = {
        "name": "Плиты ПБ 65-12-8п",
        "qty": 2,
        "price": 23643.0,
        "discount_percent": 10.0,
        "guid": "plate-guid",
    }
    offer = _offer(discount_percent=10.0, vat_amount=1.23, total_amount=1.0, lines=[line])
    document, digest = build_invoice_document(
        offer,
        action="create",
        warehouse="Склад Готовой Продукции",
        line_delivery_kopecks=[150000],
    )
    _other, other_digest = build_invoice_document(
        offer,
        action="create",
        warehouse="БСУ склад",
        line_delivery_kopecks=[150000],
    )
    undiscounted, _unused = build_invoice_document(
        _offer(
            discount_percent=0.0,
            lines=[{**line, "discount_percent": 0.0}],
        ),
        action="create",
        line_delivery_kopecks=[150000],
    )

    goods = document["Товары"]
    assert len(goods) == 1
    assert "доставк" not in goods[0]["Номенклатура"].lower()
    assert goods[0]["Цена"] == 23643.0
    assert goods[0]["ПроцентСкидки"] == 10.0
    assert goods[0]["СуммаДоставки"] == 1500.0
    assert undiscounted["Товары"][0]["СуммаДоставки"] == 1500.0
    assert document["Самовывоз"] is False
    assert document["Склад"] == "Склад Готовой Продукции"
    payable = _payable("23643", "10", 2, "1500")
    assert payable == Decimal("44057.40")
    assert document["СуммаДокумента"] == float(payable)
    assert document["СуммаНДС"] == float(_included_vat(payable))
    assert document["СуммаНДС"] != 1.23
    assert digest == other_digest


def test_invoice_snapshot_hash_changes_when_qty_changes() -> None:
    from core.invoice_export import build_invoice_document

    _same, first = build_invoice_document(_offer(), action="create")
    _again, second = build_invoice_document(_offer(), action="create")
    changed = _offer()
    changed["lines"] = [{**changed["lines"][0], "qty": 12}]
    _other, third = build_invoice_document(changed, action="create")

    assert first == second
    assert first != third
    assert len(first) == 64


def _seed_kp(
    tmp_path: Path,
    *,
    status: str = "в архиве",
    with_counterparty: bool = True,
    qty: int = 2,
    order_number: str | None = None,
    paid_at: str | None = None,
    snapshot: str | None = None,
) -> str:
    from core.kp_db_schema import init_schema

    db_path = str(tmp_path / "plita.db")
    init_schema(db_path)
    with sqlite3.connect(db_path) as conn:
        counterparty_id = None
        if with_counterparty:
            conn.execute(
                """
                INSERT INTO counterparties (
                    code_1c, guid_1c, name, name_normalized, inn, kpp,
                    is_client, is_active, source
                ) VALUES (
                    '00-00000884', 'cp-guid', 'РОМАШКА ООО', 'ромашка ооо',
                    '7604368921', '760401001', 1, 1, 'import'
                )
                """
            )
            counterparty_id = int(conn.execute("SELECT last_insert_rowid()").fetchone()[0])
        conn.execute(
            """
            INSERT INTO KP_offers (
                kp_id, creation_date, customer_name, manager_name,
                discount_percent, subtotal, vat_amount, total_amount,
                logistics_cost, counterparty_id, customer_inn, customer_kpp,
                order_number_1c, invoice_snapshot_hash
            ) VALUES (
                22, '11.09.2026', 'РОМАШКА ООО', 'Иван Иванов',
                13, 10, 1.23, 999.99,
                18600, ?, '7604368921', '760401001', ?, ?
            )
            """,
            (counterparty_id, order_number, snapshot),
        )
        conn.execute(
            """
            INSERT INTO kp_plates (
                kp_id, position_number, plate_name, qty, unit_price, discounted_price
            ) VALUES (22, 1, 'Плиты ПБ 45,4-12-8п', ?, 15421, 100)
            """,
            (qty,),
        )
        conn.execute(
            "INSERT INTO kp_meta (kp_id, status, owner_user_id, paid_at) VALUES (22, ?, 1, ?)",
            (status, paid_at),
        )
        conn.commit()
    return db_path


def _service(db_path: str, export_dir: Path):
    from app.repositories.kp_archive_repository import KpArchiveRepository
    from app.services.archive_service import ArchiveService

    return ArchiveService(
        repository=KpArchiveRepository(db_path=db_path),
        outputs_dir=export_dir.parent / "out",
        export_dir=export_dir,
    )


def _allow_guid(lines, _conn):
    return GateReport(
        ready=tuple(ReadyItem(line=line, guid="guid-line") for line in lines),
        missing=(),
    )


def _block_guid(_lines, _conn):
    raise InvoiceGuidBlockError(
        GateReport(
            ready=(),
            missing=(
                MissingItem(
                    line=OrderLine("plate", "Плиты ПБ 45,4-12-8п", 2),
                    reason="нет GUID 1С",
                    action_hint="заведите карточку в 1С и загрузите отчёт",
                ),
            ),
        )
    )


def _active_contract_row(self, counterparty_id):  # noqa: ANN001
    return {
        "counterparty_id": int(counterparty_id),
        "status": "нет",
        "number": "0001/09/26",
    }


def _no_contract(self, _counterparty_id):  # noqa: ANN001
    return None


def _status(db_path: str) -> str:
    with sqlite3.connect(db_path) as conn:
        row = conn.execute("SELECT status FROM kp_meta WHERE kp_id = 22").fetchone()
    return str(row[0])


def _warehouse(db_path: str) -> str | None:
    with sqlite3.connect(db_path) as conn:
        row = conn.execute(
            "SELECT invoice_warehouse FROM KP_offers WHERE kp_id = 22"
        ).fetchone()
    if row is None or row[0] is None:
        return None
    text = str(row[0]).strip()
    return text or None


def _snapshot(db_path: str) -> str | None:
    with sqlite3.connect(db_path) as conn:
        row = conn.execute(
            "SELECT invoice_snapshot_hash FROM KP_offers WHERE kp_id = 22"
        ).fetchone()
    return row[0]


def test_invoice_export_without_counterparty_keeps_archive_and_writes_no_file(
    tmp_path: Path,
) -> None:
    from app.services.archive_service import ArchiveValidationError

    db_path = _seed_kp(tmp_path, with_counterparty=False)
    export_dir = tmp_path / "exchange"
    service = _service(db_path, export_dir)

    with pytest.raises(ArchiveValidationError, match="контрагент"):
        service.export_invoice(22, user={"id": 1, "role": "admin"})

    assert _status(db_path) == "в архиве"
    assert not export_dir.exists() or list(export_dir.glob("*.json")) == []


def test_invoice_export_without_contract_keeps_archive(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from app.services.archive_service import ArchiveValidationError

    db_path = _seed_kp(tmp_path)
    export_dir = tmp_path / "exchange"
    monkeypatch.setattr(
        "app.services.archive_service.SupplyContractRepository.get_active_by_counterparty",
        _no_contract,
    )
    monkeypatch.setattr("app.services.archive_service.assert_invoice_guids", _allow_guid)
    service = _service(db_path, export_dir)

    with pytest.raises(ArchiveValidationError, match="договор"):
        service.export_invoice(22, user={"id": 1, "role": "admin"})

    assert _status(db_path) == "в архиве"
    assert list(export_dir.glob("*.json")) == []


def test_invoice_export_without_guid_keeps_archive(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from app.services.archive_service import ArchiveValidationError

    db_path = _seed_kp(tmp_path)
    export_dir = tmp_path / "exchange"
    monkeypatch.setattr(
        "app.services.archive_service.SupplyContractRepository.get_active_by_counterparty",
        _active_contract_row,
    )
    monkeypatch.setattr("app.services.archive_service.assert_invoice_guids", _block_guid)
    service = _service(db_path, export_dir)

    with pytest.raises(ArchiveValidationError, match="нет GUID 1С"):
        service.export_invoice(22, user={"id": 1, "role": "admin"})

    assert _status(db_path) == "в архиве"
    assert list(export_dir.glob("*.json")) == []


def test_invoice_export_writes_file_then_moves_status(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    db_path = _seed_kp(tmp_path)
    export_dir = tmp_path / "exchange"
    monkeypatch.setattr(
        "app.services.archive_service.SupplyContractRepository.get_active_by_counterparty",
        _active_contract_row,
    )
    monkeypatch.setattr("app.services.archive_service.assert_invoice_guids", _allow_guid)
    service = _service(db_path, export_dir)

    details = service.export_invoice(
        22, user={"id": 1, "role": "admin"}, warehouse=_WAREHOUSE
    )

    files = list(export_dir.glob("invoice_create_22_*.json"))
    assert len(files) == 1
    payload = json.loads(files[0].read_text(encoding="utf-8"))
    assert payload["Документ"] == "Счёт на оплату"
    assert payload["event"] == "invoice_export"
    assert payload["action"] == "create"
    assert "НомерЗаказа" not in payload
    payable = _payable("15421", "13", 2)
    assert payload["Склад"] == _WAREHOUSE
    assert payload["Самовывоз"] is True
    assert payload["Товары"][0]["СуммаДоставки"] == 0.0
    assert "Склад" not in payload["Товары"][0]
    assert _warehouse(db_path) == _WAREHOUSE
    assert payload["Товары"][0]["Цена"] == 15421.0
    assert payload["СуммаНДС"] == float(_included_vat(payable))
    assert payload["СуммаДокумента"] == float(payable)
    assert payload["СуммаНДС"] != 1.23
    keys = list(payload)
    assert keys.index("НомерДоговора") == keys.index("Контрагент") + 1
    assert payload["НомерДоговора"] == "0001/09/26"
    assert payload["Товары"][0]["GUID"] == "guid-line"
    assert details.status == "на согласовании"
    assert _status(db_path) == "на согласовании"
    assert _snapshot(db_path)


def test_invoice_export_from_on_approval_does_not_write_another_create(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from app.services.archive_service import ArchiveValidationError

    db_path = _seed_kp(tmp_path)
    export_dir = tmp_path / "exchange"
    monkeypatch.setattr(
        "app.services.archive_service.SupplyContractRepository.get_active_by_counterparty",
        _active_contract_row,
    )
    monkeypatch.setattr("app.services.archive_service.assert_invoice_guids", _allow_guid)
    service = _service(db_path, export_dir)
    service.export_invoice(22, user={"id": 1, "role": "admin"}, warehouse=_WAREHOUSE)

    with pytest.raises(ArchiveValidationError, match="в архиве"):
        service.export_invoice(22, user={"id": 1, "role": "admin"}, warehouse=_WAREHOUSE)

    assert len(list(export_dir.glob("invoice_create_22_*.json"))) == 1
    assert _status(db_path) == "на согласовании"


def test_invoice_export_file_error_does_not_change_status(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from app.services.archive_service import ArchiveError

    db_path = _seed_kp(tmp_path)
    blocked = tmp_path / "not-a-directory"
    blocked.write_text("file", encoding="utf-8")
    monkeypatch.setattr(
        "app.services.archive_service.SupplyContractRepository.get_active_by_counterparty",
        _active_contract_row,
    )
    monkeypatch.setattr("app.services.archive_service.assert_invoice_guids", _allow_guid)
    service = _service(db_path, blocked)

    with pytest.raises(ArchiveError):
        service.export_invoice(22, user={"id": 1, "role": "admin"}, warehouse=_WAREHOUSE)

    assert _status(db_path) == "в архиве"
    assert _warehouse(db_path) is None


def _patch_gates(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(
        "app.services.archive_service.SupplyContractRepository.get_active_by_counterparty",
        _active_contract_row,
    )
    monkeypatch.setattr("app.services.archive_service.assert_invoice_guids", _allow_guid)


def _set_order_number(db_path: str, number: str) -> None:
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "UPDATE KP_offers SET order_number_1c = ? WHERE kp_id = 22",
            (number,),
        )
        conn.commit()


def _set_qty(db_path: str, qty: int) -> None:
    with sqlite3.connect(db_path) as conn:
        conn.execute("UPDATE kp_plates SET qty = ? WHERE kp_id = 22", (qty,))
        conn.commit()


def test_invoice_correction_without_number_writes_no_file(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from app.services.archive_service import ArchiveValidationError

    db_path = _seed_kp(tmp_path, status="на согласовании", snapshot="stored-hash")
    export_dir = tmp_path / "exchange"
    _patch_gates(monkeypatch)
    service = _service(db_path, export_dir)

    with pytest.raises(ArchiveValidationError, match="номер"):
        service.export_invoice_correction(22, user={"id": 1, "role": "admin"})

    assert not export_dir.exists() or list(export_dir.glob("*.json")) == []
    assert _status(db_path) == "на согласовании"
    assert _snapshot(db_path) == "stored-hash"


def test_invoice_correction_same_snapshot_refuses(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from app.services.archive_service import ArchiveValidationError

    db_path = _seed_kp(tmp_path)
    export_dir = tmp_path / "exchange"
    _patch_gates(monkeypatch)
    service = _service(db_path, export_dir)
    service.export_invoice(22, user={"id": 1, "role": "admin"}, warehouse=_WAREHOUSE)
    _set_order_number(db_path, "ЯР-1")
    stored = _snapshot(db_path)

    with pytest.raises(ArchiveValidationError, match="не изменил"):
        service.export_invoice_correction(22, user={"id": 1, "role": "admin"})

    assert list(export_dir.glob("invoice_update_22_*.json")) == []
    assert _status(db_path) == "на согласовании"
    assert _snapshot(db_path) == stored


def test_invoice_correction_changed_qty_writes_update_and_keeps_status(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    db_path = _seed_kp(tmp_path)
    export_dir = tmp_path / "exchange"
    _patch_gates(monkeypatch)
    service = _service(db_path, export_dir)
    service.export_invoice(22, user={"id": 1, "role": "admin"}, warehouse=_WAREHOUSE)
    _set_order_number(db_path, "ЯР-77")
    _set_qty(db_path, 5)
    before = service.get_details(22, user={"id": 1, "role": "admin"})
    assert before.correction_pending is True
    stored = _snapshot(db_path)

    details = service.export_invoice_correction(22, user={"id": 1, "role": "admin"})

    files = list(export_dir.glob("invoice_update_22_*.json"))
    assert len(files) == 1
    payload = json.loads(files[0].read_text(encoding="utf-8"))
    assert payload["action"] == "update"
    assert payload["НомерЗаказа"] == "ЯР-77"
    assert payload["Склад"] == _WAREHOUSE
    assert "Самовывоз" in payload
    assert _warehouse(db_path) == _WAREHOUSE
    assert payload["Документ"] == "Счёт на оплату"
    assert details.status == "на согласовании"
    assert details.correction_pending is False
    assert _status(db_path) == "на согласовании"
    assert _snapshot(db_path) not in (None, stored)


def test_invoice_export_unknown_warehouse_writes_no_file(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from app.services.archive_service import ArchiveValidationError

    db_path = _seed_kp(tmp_path)
    export_dir = tmp_path / "exchange"
    _patch_gates(monkeypatch)
    service = _service(db_path, export_dir)

    with pytest.raises(ArchiveValidationError, match="Неизвестный склад"):
        service.export_invoice(22, user={"id": 1, "role": "admin"}, warehouse="Чужой склад")

    assert _status(db_path) == "в архиве"
    assert _warehouse(db_path) is None
    assert list(export_dir.glob("*.json")) == []


def test_invoice_export_empty_warehouse_writes_no_file(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from app.services.archive_service import ArchiveValidationError

    db_path = _seed_kp(tmp_path)
    export_dir = tmp_path / "exchange"
    _patch_gates(monkeypatch)
    service = _service(db_path, export_dir)

    with pytest.raises(ArchiveValidationError, match="Укажите склад"):
        service.export_invoice(22, user={"id": 1, "role": "admin"}, warehouse="  ")

    assert _status(db_path) == "в архиве"
    assert not export_dir.exists() or list(export_dir.glob("*.json")) == []


def test_invoice_export_unready_pile_boiler_writes_no_file(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from app.services.archive_service import ArchiveValidationError

    db_path = _seed_kp(tmp_path)
    with sqlite3.connect(db_path) as conn:
        conn.execute("UPDATE kp_meta SET product_type = 'piles' WHERE kp_id = 22")
        conn.execute(
            """
            INSERT INTO kp_piles (
                kp_id, position_number, mark, concrete_grade, qty, unit_price, discounted_price
            ) VALUES (22, 1, 'С99.99-9', 'B25', 2, 100, 87)
            """
        )
        conn.commit()
    export_dir = tmp_path / "exchange"
    _patch_gates(monkeypatch)
    service = _service(db_path, export_dir)

    with pytest.raises(ArchiveValidationError, match="С99.99-9"):
        service.export_invoice(22, user={"id": 1, "role": "admin"}, warehouse=_WAREHOUSE)

    assert _status(db_path) == "в архиве"
    assert _warehouse(db_path) is None
    assert list(export_dir.glob("*.json")) == []


def test_invoice_export_delivery_share_is_after_discount(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    db_path = _seed_kp(tmp_path)
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            UPDATE KP_offers
            SET logistics_cost = 1000.01, discount_percent = 10
            WHERE kp_id = 22
            """
        )
        conn.execute(
            "UPDATE kp_plates SET length_m = 6, width_m = 1.2, qty = 2 WHERE kp_id = 22"
        )
        conn.commit()
    export_dir = tmp_path / "exchange"
    _patch_gates(monkeypatch)
    service = _service(db_path, export_dir)

    service.export_invoice(22, user={"id": 1, "role": "admin"}, warehouse=_WAREHOUSE)

    payload = json.loads(
        next(export_dir.glob("invoice_create_22_*.json")).read_text(encoding="utf-8")
    )
    goods = payload["Товары"]
    assert len(goods) == 1
    assert "доставк" not in goods[0]["Номенклатура"].lower()
    assert goods[0]["Цена"] == 15421.0
    assert goods[0]["ПроцентСкидки"] == 10.0
    assert goods[0]["СуммаДоставки"] == 1000.01
    assert payload["Самовывоз"] is False
    payable = _payable("15421", "10", 2, "1000.01")
    assert payload["СуммаДокумента"] == float(payable)
    assert payload["СуммаНДС"] == float(_included_vat(payable))
    assert payload["СуммаНДС"] != 1.23


def test_invoice_correction_empty_warehouse_without_body_writes_no_file(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from app.services.archive_service import ArchiveValidationError

    db_path = _seed_kp(
        tmp_path,
        status="на согласовании",
        order_number="ЯР-3",
        snapshot="stored-hash",
    )
    export_dir = tmp_path / "exchange"
    _patch_gates(monkeypatch)
    service = _service(db_path, export_dir)

    with pytest.raises(ArchiveValidationError, match="Укажите склад"):
        service.export_invoice_correction(22, user={"id": 1, "role": "admin"})

    assert _status(db_path) == "на согласовании"
    assert _warehouse(db_path) is None
    assert _snapshot(db_path) == "stored-hash"
    assert list(export_dir.glob("*.json")) == []


def test_invoice_correction_empty_warehouse_saves_name_on_update(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    db_path = _seed_kp(
        tmp_path,
        status="на согласовании",
        order_number="ЯР-3",
        snapshot="stored-hash",
    )
    export_dir = tmp_path / "exchange"
    _patch_gates(monkeypatch)
    service = _service(db_path, export_dir)

    service.export_invoice_correction(
        22, user={"id": 1, "role": "admin"}, warehouse="БСУ склад"
    )

    files = list(export_dir.glob("invoice_update_22_*.json"))
    assert len(files) == 1
    payload = json.loads(files[0].read_text(encoding="utf-8"))
    assert payload["action"] == "update"
    assert payload["Склад"] == "БСУ склад"
    assert "Самовывоз" in payload
    assert _warehouse(db_path) == "БСУ склад"
    assert _status(db_path) == "на согласовании"
