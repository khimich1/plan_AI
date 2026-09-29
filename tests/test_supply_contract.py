"""Схема и сервис договора поставки (SC-002, SC-003)."""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest

from app.schemas.supply_contract import SupplyContractCreate
from app.services.supply_contract_service import (
    SupplyContractFieldError,
    SupplyContractService,
    SupplyContractValidationError,
)
from core import kp_db_schema
from core.kp_db_common import _connect
from core.supply_contract import CANCELLED_STATUS, MissingSupplyContractError, build_preamble

ADMIN = {"id": 7, "role": "admin", "username": "admin"}

CONTRACT_COLS = {
    "id",
    "counterparty_id",
    "imported_name",
    "number",
    "contract_date",
    "manager_name",
    "status",
    "scan_note",
    "scan_path",
    "legal_form",
    "full_name",
    "short_name",
    "signatory_position",
    "signatory_name",
    "signatory_verb",
    "authority_basis",
    "poa_number",
    "poa_date",
    "inn",
    "kpp",
    "ogrn",
    "legal_address",
    "postal_address",
    "phone",
    "email",
    "bank_name",
    "account",
    "corr_account",
    "bik",
    "edo_operator",
    "edo_id",
    "created_at",
    "created_by_user_id",
}

EVENT_COLS = {
    "id",
    "contract_id",
    "field",
    "old_value",
    "new_value",
    "user_id",
    "created_at",
}


def _fresh_db(tmp_path: Path, name: str = "supply.db") -> str:
    db_path = str(tmp_path / name)
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    return db_path


def _table_cols(conn: sqlite3.Connection, table: str) -> set[str]:
    return {row[1] for row in conn.execute(f"PRAGMA table_info({table})")}


def _seed_counterparty(conn: sqlite3.Connection, code: str = "00-1") -> int:
    cur = conn.execute(
        """
        INSERT INTO counterparties (code_1c, name, name_normalized, inn, kpp)
        VALUES (?, 'РОМАШКА ООО', 'ромашка ооо', '7701000001', '770101001')
        """,
        (code,),
    )
    return int(cur.lastrowid)


def _insert_contract(
    conn: sqlite3.Connection,
    *,
    counterparty_id: int | None,
    number: str,
    status: str = "нет",
) -> None:
    conn.execute(
        """
        INSERT INTO supply_contract (
            counterparty_id, number, contract_date, manager_name, status,
            legal_form, full_name, short_name, signatory_name, signatory_verb,
            authority_basis, inn, ogrn, legal_address, email,
            bank_name, account, corr_account, bik,
            created_at, created_by_user_id
        ) VALUES (
            ?, ?, '2026-09-25', 'Иван', ?,
            'ooo', 'Ромашка', 'Ромашка', 'Иванов', 'действующего',
            'устав', '7701000001', '1027700000001', 'Москва', 'a@b.ru',
            'Банк', '40702810000000000001', '30101810000000000001', '044525225',
            '2026-09-25T12:00:00', 1
        )
        """,
        (counterparty_id, number, status),
    )


def test_schema_creates_supply_contract_tables(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        names = {
            row[0]
            for row in conn.execute("SELECT name FROM sqlite_master WHERE type = 'table'")
        }
        assert "supply_contract" in names
        assert "supply_contract_event" in names
        assert _table_cols(conn, "supply_contract") == CONTRACT_COLS
        assert _table_cols(conn, "supply_contract_event") == EVENT_COLS
        offer_cols = _table_cols(conn, "KP_offers")
        assert "order_number_1c" in offer_cols
        assert "order_status_1c" in offer_cols


def test_schema_ensure_is_idempotent(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "idem.db")
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    kp_db_schema._init_schema_impl(db_path)
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        assert _table_cols(conn, "supply_contract") == CONTRACT_COLS


def test_schema_one_active_contract_per_counterparty(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "uniq.db")
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        client_id = _seed_counterparty(conn)
        _insert_contract(conn, counterparty_id=client_id, number="0001/09/26")
        with pytest.raises(sqlite3.IntegrityError):
            _insert_contract(conn, counterparty_id=client_id, number="0002/09/26")


def test_schema_two_cancelled_contracts_are_allowed(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "cancel.db")
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        client_id = _seed_counterparty(conn)
        _insert_contract(
            conn,
            counterparty_id=client_id,
            number="0001/09/26",
            status=CANCELLED_STATUS,
        )
        _insert_contract(
            conn,
            counterparty_id=client_id,
            number="0002/09/26",
            status=CANCELLED_STATUS,
        )
        count = conn.execute("SELECT COUNT(*) FROM supply_contract").fetchone()[0]
        assert count == 2


def test_schema_unlinked_rows_do_not_take_the_active_slot(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "import.db")
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        _insert_contract(conn, counterparty_id=None, number="0001/09/26")
        _insert_contract(conn, counterparty_id=None, number="0002/09/26")
        count = conn.execute("SELECT COUNT(*) FROM supply_contract").fetchone()[0]
        assert count == 2


def _seed_offer(
    conn: sqlite3.Connection,
    *,
    kp_id: int,
    counterparty_id: int | None,
    status: str = "в архиве",
    manager_name: str = "Иван Иванов",
) -> None:
    conn.execute(
        """
        INSERT INTO KP_offers (
            kp_id, creation_date, customer_name, manager_name, counterparty_id
        ) VALUES (?, '2026-09-01', 'РОМАШКА ООО', ?, ?)
        """,
        (kp_id, manager_name, counterparty_id),
    )
    conn.execute(
        "INSERT INTO kp_meta (kp_id, status, owner_user_id) VALUES (?, ?, 7)",
        (kp_id, status),
    )


def _buyer(**overrides: object) -> SupplyContractCreate:
    data: dict[str, object] = {
        "legal_form": "ooo",
        "full_name": "Ромашка",
        "short_name": "Ромашка",
        "signatory_position": "Генерального директора",
        "signatory_name": "Иванов Иван",
        "signatory_verb": "действующего",
        "authority_basis": "устав",
        "inn": "760 4 01001",
        "kpp": "760401001",
        "ogrn": "1027600000001",
        "legal_address": "Ярославль",
        "email": "is-ag @mail.ru",
        "bank_name": "Промсвязьбанк",
        "account": "40702 810 0 0000 0000001",
        "corr_account": "30101810000000000760",
        "bik": "044 525 555",
        "contract_date": "2026-09-25",
    }
    data.update(overrides)
    return SupplyContractCreate.model_validate(data)


def _service(db_path: str) -> SupplyContractService:
    return SupplyContractService(db_path=db_path)


def _contract_count(db_path: str) -> int:
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        return int(conn.execute("SELECT COUNT(*) FROM supply_contract").fetchone()[0])


def test_create_stores_one_row_with_status_net(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "create.db")
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        client_id = _seed_counterparty(conn)
        _seed_offer(conn, kp_id=1, counterparty_id=client_id)
        before = dict(
            conn.execute(
                "SELECT name, inn, kpp FROM counterparties WHERE id = ?",
                (client_id,),
            ).fetchone()
        )
    service = _service(db_path)

    created = service.create_for_kp(1, _buyer(), user=ADMIN)

    assert created["status"] == "нет"
    assert created["number"] == "0001/09/26"
    assert created["inn"] == "760401001"
    assert created["email"] == "is-ag@mail.ru"
    assert created["manager_name"] == "Иван Иванов"
    assert _contract_count(db_path) == 1
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        after = dict(
            conn.execute(
                "SELECT name, inn, kpp FROM counterparties WHERE id = ?",
                (client_id,),
            ).fetchone()
        )
        events = conn.execute(
            "SELECT field, old_value, new_value, user_id FROM supply_contract_event"
        ).fetchall()
    assert after == before
    assert len(events) == 1
    assert events[0]["field"] == "status"
    assert events[0]["old_value"] is None
    assert events[0]["new_value"] == "нет"
    assert events[0]["user_id"] == 7


def test_second_create_returns_same_number_without_new_seq(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "repeat.db")
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        client_id = _seed_counterparty(conn)
        _seed_offer(conn, kp_id=1, counterparty_id=client_id)
        _seed_offer(conn, kp_id=2, counterparty_id=client_id)
    service = _service(db_path)

    first = service.create_for_kp(1, _buyer(), user=ADMIN)
    second = service.create_for_kp(2, _buyer(contract_date="2026-10-01"), user=ADMIN)
    visible = service.get_for_kp(2, user=ADMIN)

    assert second["number"] == first["number"] == "0001/09/26"
    assert second["id"] == first["id"]
    assert visible["number"] == first["number"]
    assert _contract_count(db_path) == 1
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        events = conn.execute("SELECT COUNT(*) FROM supply_contract_event").fetchone()[0]
    assert events == 1


def test_ooo_requires_kpp_ip_does_not(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "kpp.db")
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        client_id = _seed_counterparty(conn)
        _seed_offer(conn, kp_id=1, counterparty_id=client_id)
    service = _service(db_path)

    with pytest.raises(SupplyContractFieldError, match="КПП"):
        service.create_for_kp(1, _buyer(kpp=None), user=ADMIN)
    assert _contract_count(db_path) == 0

    created = service.create_for_kp(
        1,
        _buyer(
            legal_form="ip",
            kpp=None,
            signatory_position=None,
            full_name="Иванов Иван Иванович",
            signatory_name="Иванов Иван Иванович",
        ),
        user=ADMIN,
    )
    assert created["kpp"] in (None, "")
    preamble = build_preamble(
        legal_form=created["legal_form"],
        full_name=created["full_name"],
        signatory_name=created["signatory_name"],
        signatory_verb=created["signatory_verb"],
        authority_basis=created["authority_basis"],
    )
    assert "Общество с ограниченной ответственностью" not in preamble


def test_date_change_keeps_the_issued_number(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "date.db")
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        client_id = _seed_counterparty(conn)
        _seed_offer(conn, kp_id=1, counterparty_id=client_id)
    service = _service(db_path)
    created = service.create_for_kp(1, _buyer(), user=ADMIN)

    updated = service.update_contract_date(
        int(created["id"]),
        "2026-11-02",
        user=ADMIN,
    )

    assert updated["number"] == created["number"] == "0001/09/26"
    assert updated["contract_date"] == "2026-11-02"
    assert _contract_count(db_path) == 1
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        event = conn.execute(
            """
            SELECT field, old_value, new_value, user_id
            FROM supply_contract_event
            WHERE field = 'contract_date'
            """
        ).fetchone()
    assert event["old_value"] == "2026-09-25"
    assert event["new_value"] == "2026-11-02"
    assert event["user_id"] == 7


def test_cancelled_number_is_skipped_for_the_next_client(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "skip.db")
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        old_id = _seed_counterparty(conn, "00-old")
        new_id = _seed_counterparty(conn, "00-new")
        _insert_contract(
            conn,
            counterparty_id=old_id,
            number="1027/09/26",
            status=CANCELLED_STATUS,
        )
        _seed_offer(conn, kp_id=3, counterparty_id=new_id)
    service = _service(db_path)

    created = service.create_for_kp(3, _buyer(), user=ADMIN)

    assert created["number"] == "1028/09/26"


def test_supply_contract_read_on_approval_status_returns_number(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "approval-read.db")
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        client_id = _seed_counterparty(conn)
        _seed_offer(conn, kp_id=1, counterparty_id=client_id)
    service = _service(db_path)
    created = service.create_for_kp(1, _buyer(), user=ADMIN)
    with _connect(db_path) as conn:
        conn.execute(
            "UPDATE kp_meta SET status = ? WHERE kp_id = 1",
            ("на согласовании",),
        )
        conn.commit()

    visible = service.get_for_kp(1, user=ADMIN)
    payload, filename = service.document_for_kp(1, user=ADMIN)

    assert visible is not None
    assert visible["number"] == created["number"]
    assert payload.startswith(b"PK")
    assert created["number"].replace("/", "-") in filename
    with pytest.raises(SupplyContractValidationError, match="в архиве"):
        service.create_for_kp(1, _buyer(), user=ADMIN)


def test_missing_counterparty_and_non_archive_are_russian_errors(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "gates.db")
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        client_id = _seed_counterparty(conn)
        _seed_offer(conn, kp_id=1, counterparty_id=None)
        _seed_offer(conn, kp_id=2, counterparty_id=client_id, status="в работе")
    service = _service(db_path)

    with pytest.raises(SupplyContractValidationError, match="Сначала занесите контрагента из 1С"):
        service.create_for_kp(1, _buyer(), user=ADMIN)
    with pytest.raises(SupplyContractValidationError, match="в архиве"):
        service.create_for_kp(2, _buyer(), user=ADMIN)
    assert _contract_count(db_path) == 0


def test_registry_status_edo_keeps_number_and_writes_event(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "edo.db")
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        client_id = _seed_counterparty(conn)
        _seed_offer(conn, kp_id=1, counterparty_id=client_id)
    service = _service(db_path)
    created = service.create_for_kp(1, _buyer(), user=ADMIN)

    updated = service.update_contract(
        int(created["id"]),
        user=ADMIN,
        status="подписан по ЭДО",
    )

    assert updated["number"] == created["number"] == "0001/09/26"
    assert updated["status"] == "подписан по ЭДО"
    assert "scan_path" not in updated
    assert service.get_for_kp(1, user=ADMIN)["number"] == "0001/09/26"
    assert _contract_count(db_path) == 1
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        event = conn.execute(
            """
            SELECT field, old_value, new_value, user_id
            FROM supply_contract_event
            WHERE field = 'status' AND new_value = 'подписан по ЭДО'
            """
        ).fetchone()
    assert event["old_value"] == "нет"
    assert event["new_value"] == "подписан по ЭДО"
    assert event["user_id"] == 7


def test_registry_cancel_drops_active_contract(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "cancel-active.db")
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        client_id = _seed_counterparty(conn)
        _seed_offer(conn, kp_id=1, counterparty_id=client_id)
    service = _service(db_path)
    created = service.create_for_kp(1, _buyer(), user=ADMIN)

    service.update_contract(int(created["id"]), user=ADMIN, status=CANCELLED_STATUS)

    assert service.get_for_kp(1, user=ADMIN) is None
    assert _contract_count(db_path) == 1
    rows = service.list_registry(user=ADMIN)
    assert rows[0]["number"] == "0001/09/26"
    assert rows[0]["status"] == CANCELLED_STATUS
    assert rows[0]["counterparty_name"] == "РОМАШКА ООО"
    assert "scan_path" not in rows[0]


def test_registry_scan_file_is_on_disk_and_hidden_from_payload(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "scan.db")
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        client_id = _seed_counterparty(conn)
        _seed_offer(conn, kp_id=1, counterparty_id=client_id)
    service = SupplyContractService(db_path=db_path, scans_dir=tmp_path / "scans")
    created = service.create_for_kp(1, _buyer(), user=ADMIN)

    updated = service.update_contract(
        int(created["id"]),
        user=ADMIN,
        scan_note="получен скан",
        scan_note_set=True,
        scan_bytes=b"%PDF-1.4 scan",
        scan_filename="card.pdf",
    )

    assert updated["has_scan"] is True
    assert updated["scan_note"] == "получен скан"
    assert "scan_path" not in updated
    stored = list((tmp_path / "scans").rglob("*.pdf"))
    assert len(stored) == 1
    assert stored[0].read_bytes().startswith(b"%PDF")
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        event = conn.execute(
            "SELECT field, old_value, new_value, user_id FROM supply_contract_event WHERE field = 'scan_note'"
        ).fetchone()
    assert event["old_value"] is None
    assert event["new_value"] == "получен скан"
    assert event["user_id"] == 7


def test_registry_date_patch_does_not_renumber(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "date-patch.db")
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        client_id = _seed_counterparty(conn)
        _seed_offer(conn, kp_id=1, counterparty_id=client_id)
    service = _service(db_path)
    created = service.create_for_kp(1, _buyer(), user=ADMIN)

    updated = service.update_contract(
        int(created["id"]),
        user=ADMIN,
        contract_date="2026-12-01",
    )

    assert updated["number"] == "0001/09/26"
    assert updated["contract_date"] == "2026-12-01"
    assert _contract_count(db_path) == 1


def test_replace_without_cancel_rejects_and_does_not_insert(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "replace-no.db")
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        client_id = _seed_counterparty(conn)
        _seed_offer(conn, kp_id=1, counterparty_id=client_id)
    service = _service(db_path)
    created = service.create_for_kp(1, _buyer(), user=ADMIN)

    with pytest.raises(SupplyContractValidationError, match="У контрагента уже есть договор 0001/09/26"):
        service.replace_contract(int(created["id"]), _buyer(), user=ADMIN)

    assert _contract_count(db_path) == 1


def test_replace_after_cancel_issues_new_seq_and_keeps_old_number(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "replace-yes.db")
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        client_id = _seed_counterparty(conn)
        _seed_offer(conn, kp_id=1, counterparty_id=client_id)
    service = _service(db_path)
    created = service.create_for_kp(1, _buyer(), user=ADMIN)
    service.update_contract(int(created["id"]), user=ADMIN, status=CANCELLED_STATUS)

    replacement = service.replace_contract(int(created["id"]), _buyer(contract_date="2026-10-02"), user=ADMIN)

    assert replacement["number"] == "0002/10/26"
    assert replacement["status"] == "нет"
    assert _contract_count(db_path) == 2
    numbers = {row["number"] for row in service.list_registry(user=ADMIN)}
    assert numbers == {"0001/09/26", "0002/10/26"}


def _kp_document() -> dict:
    return {
        "Документ": "КоммерческоеПредложение",
        "Контрагент": {"Наименование": "РОМАШКА ООО", "ИНН": "7701000001"},
        "Товары": [{"Номенклатура": "Плиты ПБ", "Количество": 2}],
    }


def test_stamp_accepts_invoice_and_commercial_offer_document_names() -> None:
    from core.supply_contract import KP_DOCUMENT, attach_contract_number as stamp

    contract = {"counterparty_id": 7, "status": "нет", "number": "0001/09/26"}
    for document_name in ("Счёт на оплату", KP_DOCUMENT):
        source = {
            "Документ": document_name,
            "Контрагент": {"Наименование": "РОМАШКА ООО"},
            "Товары": [{"Номенклатура": "Плиты ПБ", "Количество": 2}],
        }
        stamped = stamp(source, 7, contract)
        keys = list(stamped)
        assert stamped["НомерДоговора"] == "0001/09/26"
        assert keys.index("НомерДоговора") == keys.index("Контрагент") + 1
        assert stamped["Документ"] == document_name
        assert "НомерДоговора" not in source


def test_json_header_puts_active_number_beside_counterparty(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "json-header.db")
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        client_id = _seed_counterparty(conn)
        _seed_offer(conn, kp_id=1, counterparty_id=client_id)
    service = _service(db_path)
    created = service.create_for_kp(1, _buyer(), user=ADMIN)
    source = _kp_document()

    stamped = service.attach_contract_number(source, client_id)

    assert stamped["НомерДоговора"] == created["number"]
    keys = list(stamped)
    assert keys.index("НомерДоговора") == keys.index("Контрагент") + 1
    assert "НомерДоговора" not in stamped["Контрагент"]
    assert stamped["Товары"] == source["Товары"]
    assert all("НомерДоговора" not in line for line in stamped["Товары"])
    assert "НомерДоговора" not in source


def test_json_header_without_active_contract_raises(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "json-missing.db")
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        client_id = _seed_counterparty(conn)
        _insert_contract(conn, counterparty_id=client_id, number="0001/09/26", status=CANCELLED_STATUS)
    source = _kp_document()

    with pytest.raises(MissingSupplyContractError, match="Нет действующего договора"):
        _service(db_path).attach_contract_number(source, client_id)

    assert "НомерДоговора" not in source
    assert source["Товары"][0] == {"Номенклатура": "Плиты ПБ", "Количество": 2}
