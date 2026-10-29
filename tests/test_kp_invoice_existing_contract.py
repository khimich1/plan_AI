"""Связь договора с единственным активным клиентом по ИНН перед счётом в 1С."""

from __future__ import annotations

import json
import sqlite3
from pathlib import Path

import pytest

from core.supply_contract import CANCELLED_STATUS

KARAVAN_INN = "3528341614"


def _contract(**overrides: object) -> dict:
    row: dict = {
        "id": 10,
        "counterparty_id": None,
        "status": "нет",
        "inn": KARAVAN_INN,
        "contract_date": "2025-09-08",
        "number": "0882/09/25",
    }
    row.update(overrides)
    return row


def _client(**overrides: object) -> dict:
    row: dict = {
        "id": 1362,
        "inn": KARAVAN_INN,
        "is_active": 1,
        "is_client": 1,
        "code_1c": "00-00003032",
        "name": "КАРАВАН ООО",
    }
    row.update(overrides)
    return row


def _links(contracts: list[dict], clients: list[dict]) -> list[tuple[int, int]]:
    from app.repositories.contract_client_link import select_contract_client_links

    return select_contract_client_links(contracts, clients)


def test_one_contract_and_one_active_client_link_karavan() -> None:
    pairs = _links([_contract()], [_client()])

    assert pairs == [(10, 1362)]


def test_two_contracts_of_one_inn_link_nothing() -> None:
    early = _contract(id=1, number="0100/01/24", contract_date="2024-01-15")
    late = _contract(id=2, number="0882/09/25", contract_date="2025-09-08")

    assert _links([early, late], [_client()]) == []


def test_two_contracts_with_the_same_date_link_nothing() -> None:
    smaller = _contract(id=4, number="0004/09/25", contract_date="2025-09-08")
    larger = _contract(id=9, number="0009/09/25", contract_date="2025-09-08")

    assert _links([smaller, larger], [_client()]) == []


def test_two_directory_rows_for_one_inn_link_nothing() -> None:
    clients = [
        _client(id=1362, code_1c="00-00003032"),
        _client(id=2000, code_1c="00-00009999", name="КАРАВАН ДУБЛЬ"),
    ]

    assert _links([_contract()], clients) == []


def test_inactive_extra_row_still_makes_inn_ambiguous() -> None:
    clients = [
        _client(id=1362),
        _client(id=2000, code_1c="00-00009999", is_active=0),
    ]

    assert _links([_contract()], clients) == []


def test_inn_missing_from_directory_links_nothing() -> None:
    assert _links([_contract()], []) == []


def test_single_inactive_or_non_client_row_links_nothing() -> None:
    inactive = [_client(is_active=0)]
    supplier = [_client(is_client=0)]

    assert _links([_contract()], inactive) == []
    assert _links([_contract()], supplier) == []


def test_cancelled_and_already_linked_contracts_are_not_chosen() -> None:
    cancelled = _contract(id=1, number="0001/01/26", status=CANCELLED_STATUS, contract_date="2026-10-01")
    linked = _contract(
        id=2,
        number="0001/10/26",
        counterparty_id=50,
        contract_date="2026-10-02",
    )
    free = _contract(id=3, number="0882/09/25", contract_date="2025-09-08")

    pairs = _links([cancelled, linked, free], [_client()])

    assert pairs == [(3, 1362)]
    assert linked["counterparty_id"] == 50


def test_client_with_another_active_contract_links_nothing() -> None:
    owned = _contract(id=1, number="0001/10/26", counterparty_id=1362, contract_date="2026-10-01")
    other = _contract(id=2, number="0882/09/25", contract_date="2026-10-02")

    assert _links([owned, other], [_client()]) == []


def test_cancelled_contract_on_client_does_not_block_a_new_link() -> None:
    cancelled = _contract(
        id=1,
        number="0001/01/20",
        counterparty_id=1362,
        status=CANCELLED_STATUS,
        contract_date="2020-01-01",
    )
    fresh = _contract(id=8, number="0882/09/25", contract_date="2025-09-08")

    assert _links([cancelled, fresh], [_client()]) == [(8, 1362)]


def test_latest_cancelled_contract_is_skipped_for_an_earlier_active_one() -> None:
    cancelled = _contract(
        id=7,
        number="0999/10/26",
        status=CANCELLED_STATUS,
        contract_date="2026-10-20",
    )
    active = _contract(id=3, number="0882/09/25", contract_date="2025-09-08")

    assert _links([cancelled, active], [_client()]) == [(3, 1362)]


def test_blank_inn_does_not_match_another_blank_inn() -> None:
    contracts = [_contract(inn=None), _contract(id=11, inn="  ", number="0002/01/26")]
    clients = [_client(inn=None), _client(id=12, inn="", code_1c="00-00000002")]

    assert _links(contracts, clients) == []


def test_inn_whitespace_still_matches_one_active_client() -> None:
    pairs = _links(
        [_contract(inn=" 3528341614 ")],
        [_client(inn="3528341614")],
    )

    assert pairs == [(10, 1362)]


def test_dated_and_undated_contracts_of_one_inn_link_nothing() -> None:
    undated = _contract(id=2, number="0002/01/26", contract_date=None)
    dated = _contract(id=1, number="0882/09/25", contract_date="2025-09-08")

    assert _links([undated, dated], [_client()]) == []


def test_two_unique_inns_produce_two_pairs() -> None:
    contracts = [
        _contract(id=10, inn=KARAVAN_INN, number="0882/09/25"),
        _contract(id=20, inn="1111111111", number="0100/01/25", contract_date="2025-01-01"),
    ]
    clients = [
        _client(id=1362, inn=KARAVAN_INN),
        _client(id=77, inn="1111111111", code_1c="00-00000077", name="ДРУГОЙ"),
    ]

    assert _links(contracts, clients) == [(10, 1362), (20, 77)]


def _fresh_db(tmp_path: Path) -> str:
    from core.kp_db_schema import init_schema

    db_path = str(tmp_path / "plita.db")
    init_schema(db_path)
    return db_path


def _insert_client(
    conn: sqlite3.Connection,
    *,
    inn: str,
    code: str,
    name: str,
    is_client: int = 1,
    is_active: int = 1,
    guid: str | None = None,
) -> int:
    conn.execute(
        """
        INSERT INTO counterparties (
            code_1c, guid_1c, name, name_normalized, inn, kpp,
            is_client, is_active, source
        ) VALUES (?, ?, ?, ?, ?, '352801001', ?, ?, 'import')
        """,
        (code, guid, name, name.casefold(), inn, is_client, is_active),
    )
    return int(conn.execute("SELECT last_insert_rowid()").fetchone()[0])


def _insert_supply_contract(
    conn: sqlite3.Connection,
    *,
    number: str,
    inn: str | None,
    contract_date: str | None,
    counterparty_id: int | None = None,
    status: str = "нет",
) -> int:
    conn.execute(
        """
        INSERT INTO supply_contract (
            counterparty_id, number, contract_date, status, inn,
            created_at, created_by_user_id
        ) VALUES (?, ?, ?, ?, ?, '2026-10-21T00:00:00', 1)
        """,
        (counterparty_id, number, contract_date, status, inn),
    )
    return int(conn.execute("SELECT last_insert_rowid()").fetchone()[0])


def _contract_owner(db_path: str, number: str) -> int | None:
    with sqlite3.connect(db_path) as conn:
        row = conn.execute(
            "SELECT counterparty_id FROM supply_contract WHERE number = ?",
            (number,),
        ).fetchone()
    return None if row[0] is None else int(row[0])


def _counterparty_count(db_path: str) -> int:
    with sqlite3.connect(db_path) as conn:
        return int(conn.execute("SELECT COUNT(*) FROM counterparties").fetchone()[0])


def test_link_writes_nothing_when_two_contracts_share_an_inn(tmp_path: Path) -> None:
    from app.repositories.supply_contract_repository import SupplyContractRepository

    db_path = _fresh_db(tmp_path)
    with sqlite3.connect(db_path) as conn:
        _insert_client(
            conn, inn=KARAVAN_INN, code="00-00003032", name="КАРАВАН ООО"
        )
        _insert_supply_contract(
            conn, number="0100/01/24", inn=KARAVAN_INN, contract_date="2024-01-15"
        )
        _insert_supply_contract(
            conn, number="0882/09/25", inn=KARAVAN_INN, contract_date="2025-09-08"
        )
        conn.commit()
    before = _counterparty_count(db_path)

    stats = SupplyContractRepository(db_path=db_path).link_unique_inn_contracts(user_id=1)

    assert stats.created == 0
    assert _contract_owner(db_path, "0882/09/25") is None
    assert _contract_owner(db_path, "0100/01/24") is None
    assert _counterparty_count(db_path) == before


def test_second_link_pass_creates_nothing_after_the_only_contract_is_linked(
    tmp_path: Path,
) -> None:
    from app.repositories.supply_contract_repository import SupplyContractRepository

    db_path = _fresh_db(tmp_path)
    with sqlite3.connect(db_path) as conn:
        client_id = _insert_client(
            conn, inn=KARAVAN_INN, code="00-00003032", name="КАРАВАН ООО"
        )
        _insert_supply_contract(
            conn, number="0882/09/25", inn=KARAVAN_INN, contract_date="2025-09-08"
        )
        conn.commit()
    repo = SupplyContractRepository(db_path=db_path)

    first = repo.link_unique_inn_contracts(user_id=1)
    second = repo.link_unique_inn_contracts(user_id=1)

    assert first.created == 1
    assert second.created == 0
    assert _contract_owner(db_path, "0882/09/25") == client_id


def test_link_does_not_rewrite_0001_or_fill_a_second_contract(tmp_path: Path) -> None:
    from app.repositories.supply_contract_repository import SupplyContractRepository

    db_path = _fresh_db(tmp_path)
    with sqlite3.connect(db_path) as conn:
        kept = _insert_client(conn, inn="5001000001", code="00-00000001", name="УЖЕ СВЯЗАН")
        _insert_supply_contract(
            conn,
            number="0001/10/26",
            inn="5001000001",
            contract_date="2026-10-01",
            counterparty_id=kept,
        )
        _insert_supply_contract(
            conn, number="0002/10/26", inn="5001000001", contract_date="2026-10-20"
        )
        conn.commit()

    stats = SupplyContractRepository(db_path=db_path).link_unique_inn_contracts()

    assert stats.created == 0
    assert _contract_owner(db_path, "0001/10/26") == kept
    assert _contract_owner(db_path, "0002/10/26") is None


def test_link_skips_ambiguous_inn_cancelled_and_unknown_inn(tmp_path: Path) -> None:
    from app.repositories.supply_contract_repository import SupplyContractRepository

    db_path = _fresh_db(tmp_path)
    with sqlite3.connect(db_path) as conn:
        _insert_client(conn, inn="3528341614", code="00-00003032", name="КАРАВАН ООО")
        _insert_client(conn, inn="3528341614", code="00-00003033", name="КАРАВАН ДУБЛЬ")
        _insert_client(conn, inn="7700000000", code="00-00000010", name="ОДИН")
        _insert_supply_contract(
            conn, number="0882/09/25", inn="3528341614", contract_date="2025-09-08"
        )
        _insert_supply_contract(
            conn,
            number="0999/10/26",
            inn="7700000000",
            contract_date="2026-10-20",
            status=CANCELLED_STATUS,
        )
        _insert_supply_contract(
            conn, number="0300/01/25", inn="9999999999", contract_date="2025-01-01"
        )
        conn.commit()
    before = _counterparty_count(db_path)

    stats = SupplyContractRepository(db_path=db_path).link_unique_inn_contracts()

    assert stats.created == 0
    assert _contract_owner(db_path, "0882/09/25") is None
    assert _contract_owner(db_path, "0999/10/26") is None
    assert _contract_owner(db_path, "0300/01/25") is None
    assert _counterparty_count(db_path) == before


_WAREHOUSE = "Склад Готовой Продукции"
_ADMIN = {"id": 1, "role": "admin"}
_SPEC_JSON = (
    '{"payment":"prepay_100","term":"by_date","term_date":"2026-10-01",'
    '"delivery":"pickup","spec_date":"2026-09-18","composition_hash":"x"}'
)


def _allow_guid(lines, _conn):
    from core.guid_gate import GateReport, ReadyItem

    return GateReport(
        ready=tuple(ReadyItem(line=line, guid="guid-line") for line in lines),
        missing=(),
    )


def _invoice_service(db_path: str, export_dir: Path):
    from app.repositories.kp_archive_repository import KpArchiveRepository
    from app.services.archive_service import ArchiveService

    return ArchiveService(
        repository=KpArchiveRepository(db_path=db_path),
        outputs_dir=export_dir.parent / "out",
        export_dir=export_dir,
    )


def _seed_invoice_kp(
    tmp_path: Path,
    *,
    counterparty_id: int | None,
    customer_name: str = "КАРАВАН на карточке",
    customer_inn: str = KARAVAN_INN,
    plate_name: str | None = "Плиты ПБ 45,4-12-8п",
) -> str:
    db_path = _fresh_db(tmp_path)
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            INSERT INTO KP_offers (
                kp_id, creation_date, customer_name, manager_name,
                discount_percent, subtotal, vat_amount, total_amount,
                logistics_cost, counterparty_id, customer_inn, customer_kpp
            ) VALUES (
                15, '11.09.2026', ?, 'Иван Иванов',
                0, 10, 0, 10,
                0, ?, ?, '352801001'
            )
            """,
            (customer_name, counterparty_id, customer_inn),
        )
        if plate_name is not None:
            conn.execute(
                """
                INSERT INTO kp_plates (
                    kp_id, position_number, plate_name, qty, unit_price, discounted_price
                ) VALUES (15, 1, ?, 2, 15421, 15421)
                """,
                (plate_name,),
            )
        conn.execute(
            "INSERT INTO kp_meta (kp_id, status, owner_user_id) VALUES (15, 'в архиве', 1)"
        )
        conn.commit()
    return db_path


def _kp_snapshot(db_path: str) -> tuple:
    with sqlite3.connect(db_path) as conn:
        return conn.execute(
            """
            SELECT customer_name, customer_inn, counterparty_id
            FROM KP_offers WHERE kp_id = 15
            """
        ).fetchone()


def _invoice_payload(export_dir: Path) -> dict:
    files = list(export_dir.glob("new_15_*.json"))
    assert len(files) == 1
    return json.loads(files[0].read_text(encoding="utf-8"))


def test_export_leaves_unlinked_contract_and_does_not_build_document(
    tmp_path: Path,
) -> None:
    from app.services.archive_service import ArchiveValidationError, MSG_NEED_CONTRACT

    db_path = _seed_invoice_kp(tmp_path, counterparty_id=None)
    with sqlite3.connect(db_path) as conn:
        client_id = _insert_client(
            conn,
            inn=KARAVAN_INN,
            code="00-00003032",
            name="КАРАВАН ООО",
            guid=None,
        )
        conn.execute(
            "UPDATE KP_offers SET counterparty_id = ? WHERE kp_id = 15",
            (client_id,),
        )
        _insert_supply_contract(
            conn, number="0882/09/25", inn=KARAVAN_INN, contract_date="2025-09-08"
        )
        conn.commit()
    before = _counterparty_count(db_path)
    export_dir = tmp_path / "exchange"
    service = _invoice_service(db_path, export_dir)

    details = service.get_details(15, user=_ADMIN)
    assert details.invoice_export_block == MSG_NEED_CONTRACT
    with pytest.raises(ArchiveValidationError, match="Нет действующего договора"):
        service.export_invoice(15, user=_ADMIN, warehouse=_WAREHOUSE)

    assert _contract_owner(db_path, "0882/09/25") is None
    assert _counterparty_count(db_path) == before
    assert list(export_dir.glob("*.json")) == []


def test_export_stamps_already_linked_contract_and_keeps_empty_guid(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from app.services.archive_service import MSG_NEED_CONTRACT

    db_path = _seed_invoice_kp(tmp_path, counterparty_id=None)
    with sqlite3.connect(db_path) as conn:
        client_id = _insert_client(
            conn,
            inn=KARAVAN_INN,
            code="00-00003032",
            name="КАРАВАН ООО",
            guid=None,
        )
        conn.execute(
            "UPDATE KP_offers SET counterparty_id = ? WHERE kp_id = 15",
            (client_id,),
        )
        _insert_supply_contract(
            conn,
            number="0882/09/25",
            inn=KARAVAN_INN,
            contract_date="2025-09-08",
            counterparty_id=client_id,
        )
        conn.commit()
    before = _counterparty_count(db_path)
    export_dir = tmp_path / "exchange"
    monkeypatch.setattr("app.services.archive_service.assert_invoice_guids", _allow_guid)
    service = _invoice_service(db_path, export_dir)

    details = service.get_details(15, user=_ADMIN)
    assert details.invoice_export_block != MSG_NEED_CONTRACT

    service.export_invoice(15, user=_ADMIN, warehouse=_WAREHOUSE)

    payload = _invoice_payload(export_dir)
    name, inn, linked = _kp_snapshot(db_path)
    assert payload["НомерДоговора"] == "0882/09/25"
    assert payload["Контрагент"]["Код"] == "00-00003032"
    assert payload["Контрагент"]["GUID"] is None
    assert _contract_owner(db_path, "0882/09/25") == client_id
    assert _counterparty_count(db_path) == before
    assert name == "КАРАВАН на карточке"
    assert inn == KARAVAN_INN
    assert linked == client_id


def test_export_refuses_when_two_contracts_share_an_inn(tmp_path: Path) -> None:
    from app.services.archive_service import ArchiveValidationError, MSG_NEED_CONTRACT

    db_path = _seed_invoice_kp(tmp_path, counterparty_id=None)
    with sqlite3.connect(db_path) as conn:
        client_id = _insert_client(
            conn, inn=KARAVAN_INN, code="00-00003032", name="КАРАВАН ООО"
        )
        conn.execute(
            "UPDATE KP_offers SET counterparty_id = ? WHERE kp_id = 15",
            (client_id,),
        )
        _insert_supply_contract(
            conn, number="0100/01/24", inn=KARAVAN_INN, contract_date="2024-01-15"
        )
        _insert_supply_contract(
            conn, number="0882/09/25", inn=KARAVAN_INN, contract_date="2025-09-08"
        )
        conn.commit()
    before = _counterparty_count(db_path)
    export_dir = tmp_path / "exchange"
    service = _invoice_service(db_path, export_dir)

    details = service.get_details(15, user=_ADMIN)
    assert details.invoice_export_block == MSG_NEED_CONTRACT
    with pytest.raises(ArchiveValidationError, match="Нет действующего договора"):
        service.export_invoice(15, user=_ADMIN, warehouse=_WAREHOUSE)

    assert _contract_owner(db_path, "0100/01/24") is None
    assert _contract_owner(db_path, "0882/09/25") is None
    assert _counterparty_count(db_path) == before
    assert list(export_dir.glob("*.json")) == []


def test_two_directory_rows_still_report_missing_contract(tmp_path: Path) -> None:
    from app.services.archive_service import ArchiveValidationError, MSG_NEED_CONTRACT

    db_path = _seed_invoice_kp(tmp_path, counterparty_id=None)
    with sqlite3.connect(db_path) as conn:
        first = _insert_client(
            conn, inn=KARAVAN_INN, code="00-00003032", name="КАРАВАН ООО"
        )
        _insert_client(conn, inn=KARAVAN_INN, code="00-00003033", name="КАРАВАН ДУБЛЬ")
        conn.execute(
            "UPDATE KP_offers SET counterparty_id = ? WHERE kp_id = 15",
            (first,),
        )
        _insert_supply_contract(
            conn, number="0882/09/25", inn=KARAVAN_INN, contract_date="2025-09-08"
        )
        conn.commit()
    before = _counterparty_count(db_path)
    export_dir = tmp_path / "exchange"
    service = _invoice_service(db_path, export_dir)

    details = service.get_details(15, user=_ADMIN)
    assert details.invoice_export_block == MSG_NEED_CONTRACT
    with pytest.raises(ArchiveValidationError, match="Нет действующего договора"):
        service.export_invoice(15, user=_ADMIN, warehouse=_WAREHOUSE)

    assert _contract_owner(db_path, "0882/09/25") is None
    assert _counterparty_count(db_path) == before
    assert list(export_dir.glob("*.json")) == []


def test_kp_without_counterparty_remembers_client_who_already_has_a_contract(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    db_path = _seed_invoice_kp(
        tmp_path,
        counterparty_id=None,
        customer_name="Имя с карточки",
        customer_inn=KARAVAN_INN,
    )
    with sqlite3.connect(db_path) as conn:
        client_id = _insert_client(
            conn, inn=KARAVAN_INN, code="00-00003032", name="КАРАВАН ООО"
        )
        _insert_supply_contract(
            conn,
            number="0882/09/25",
            inn=KARAVAN_INN,
            contract_date="2025-09-08",
            counterparty_id=client_id,
        )
        conn.commit()
    before = _counterparty_count(db_path)
    export_dir = tmp_path / "exchange"
    monkeypatch.setattr("app.services.archive_service.assert_invoice_guids", _allow_guid)

    _invoice_service(db_path, export_dir).export_invoice(
        15, user=_ADMIN, warehouse=_WAREHOUSE
    )

    name, inn, linked = _kp_snapshot(db_path)
    assert linked == client_id
    assert name == "Имя с карточки"
    assert inn == KARAVAN_INN
    assert _invoice_payload(export_dir)["НомерДоговора"] == "0882/09/25"
    assert _invoice_payload(export_dir)["Контрагент"]["Код"] == "00-00003032"
    assert _contract_owner(db_path, "0882/09/25") == client_id
    assert _counterparty_count(db_path) == before


def test_kp_without_counterparty_does_not_link_a_free_contract(tmp_path: Path) -> None:
    from app.services.archive_service import ArchiveValidationError

    db_path = _seed_invoice_kp(
        tmp_path,
        counterparty_id=None,
        customer_name="Имя с карточки",
        customer_inn=KARAVAN_INN,
    )
    with sqlite3.connect(db_path) as conn:
        _insert_client(conn, inn=KARAVAN_INN, code="00-00003032", name="КАРАВАН ООО")
        _insert_supply_contract(
            conn, number="0882/09/25", inn=KARAVAN_INN, contract_date="2025-09-08"
        )
        conn.commit()
    before = _counterparty_count(db_path)
    export_dir = tmp_path / "exchange"

    with pytest.raises(ArchiveValidationError, match="контрагент"):
        _invoice_service(db_path, export_dir).export_invoice(
            15, user=_ADMIN, warehouse=_WAREHOUSE
        )

    _name, _inn, linked = _kp_snapshot(db_path)
    assert linked is None
    assert _contract_owner(db_path, "0882/09/25") is None
    assert _counterparty_count(db_path) == before
    assert list(export_dir.glob("*.json")) == []


def test_export_does_not_use_another_clients_contract(tmp_path: Path) -> None:
    from app.services.archive_service import ArchiveValidationError, MSG_NEED_CONTRACT

    db_path = _seed_invoice_kp(tmp_path, counterparty_id=None)
    with sqlite3.connect(db_path) as conn:
        karavan = _insert_client(
            conn, inn=KARAVAN_INN, code="00-00003032", name="КАРАВАН ООО"
        )
        other = _insert_client(
            conn, inn="5001000001", code="00-00000001", name="УЖЕ СВЯЗАН"
        )
        conn.execute(
            "UPDATE KP_offers SET counterparty_id = ? WHERE kp_id = 15",
            (karavan,),
        )
        _insert_supply_contract(
            conn,
            number="0001/10/26",
            inn="5001000001",
            contract_date="2026-10-01",
            counterparty_id=other,
        )
        _insert_supply_contract(
            conn, number="0882/09/25", inn=KARAVAN_INN, contract_date="2025-09-08"
        )
        conn.commit()
    before = _counterparty_count(db_path)
    export_dir = tmp_path / "exchange"
    service = _invoice_service(db_path, export_dir)

    assert service.get_details(15, user=_ADMIN).invoice_export_block == MSG_NEED_CONTRACT
    with pytest.raises(ArchiveValidationError, match="Нет действующего договора"):
        service.export_invoice(15, user=_ADMIN, warehouse=_WAREHOUSE)

    assert _contract_owner(db_path, "0001/10/26") == other
    assert _contract_owner(db_path, "0882/09/25") is None
    assert _counterparty_count(db_path) == before
    assert list(export_dir.glob("*.json")) == []


def test_mark_without_guid_is_refused_as_guid_not_missing_contract(
    tmp_path: Path,
) -> None:
    from app.services.archive_service import ArchiveValidationError

    db_path = _seed_invoice_kp(
        tmp_path,
        counterparty_id=None,
        plate_name="МАРКА-КОТОРОЙ-НЕТ-INV003",
    )
    with sqlite3.connect(db_path) as conn:
        client_id = _insert_client(
            conn, inn=KARAVAN_INN, code="00-00003032", name="КАРАВАН ООО"
        )
        conn.execute(
            "UPDATE KP_offers SET counterparty_id = ? WHERE kp_id = 15",
            (client_id,),
        )
        _insert_supply_contract(
            conn,
            number="0882/09/25",
            inn=KARAVAN_INN,
            contract_date="2025-09-08",
            counterparty_id=client_id,
        )
        conn.commit()
    export_dir = tmp_path / "exchange"

    with pytest.raises(ArchiveValidationError) as caught:
        _invoice_service(db_path, export_dir).export_invoice(
            15, user=_ADMIN, warehouse=_WAREHOUSE
        )

    text = str(caught.value)
    assert "GUID" in text
    assert "Нет действующего договора" not in text
    assert _contract_owner(db_path, "0882/09/25") == client_id
    assert list(export_dir.glob("*.json")) == []


def test_offer_without_marks_still_stamps_the_contract_number(
    tmp_path: Path,
) -> None:
    db_path = _seed_invoice_kp(tmp_path, counterparty_id=None, plate_name=None)
    with sqlite3.connect(db_path) as conn:
        client_id = _insert_client(
            conn, inn=KARAVAN_INN, code="00-00003032", name="КАРАВАН ООО"
        )
        conn.execute(
            "UPDATE KP_offers SET counterparty_id = ? WHERE kp_id = 15",
            (client_id,),
        )
        _insert_supply_contract(
            conn,
            number="0882/09/25",
            inn=KARAVAN_INN,
            contract_date="2025-09-08",
            counterparty_id=client_id,
        )
        conn.commit()
    export_dir = tmp_path / "exchange"

    _invoice_service(db_path, export_dir).export_invoice(
        15, user=_ADMIN, warehouse=_WAREHOUSE
    )

    assert _invoice_payload(export_dir)["НомерДоговора"] == "0882/09/25"
    assert _invoice_payload(export_dir)["Товары"] == []
    assert _contract_owner(db_path, "0882/09/25") == client_id


def test_move_to_production_does_not_link_contract(tmp_path: Path) -> None:
    from datetime import date

    from app.repositories.kp_archive_repository import KpArchiveRepository
    from app.services.archive_service import ArchiveService
    from app.services.promise_service import PromiseService
    from tests.helpers.kp_db_fixtures import (
        make_iso_db,
        seed_kp_offer,
        seed_plate,
        seed_test_counterparty,
    )

    db_path = make_iso_db(tmp_path)
    seed_kp_offer(db_path, 20, status="в архиве", customer_name="Секлюцкий")
    seed_plate(
        db_path,
        kp_id=20,
        plate_name="ПБ 60-12-8",
        length_m=6.0,
        width_m=1.2,
        qty=4,
        status="в производстве",
    )
    client_id = seed_test_counterparty(db_path)
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "UPDATE KP_offers SET counterparty_id = ?, customer_inn = ? WHERE kp_id = 20",
            (client_id, "7701000001"),
        )
        _insert_supply_contract(
            conn, number="0882/09/25", inn="7701000001", contract_date="2025-09-08"
        )
        conn.commit()
    service = ArchiveService(
        repository=KpArchiveRepository(db_path=db_path),
        outputs_dir=tmp_path / "out",
        promise_service=PromiseService(
            db_path=db_path,
            occupancy_loader=lambda: {},
            today=date(2026, 9, 2),
            is_workday=lambda day: day.weekday() < 5,
        ),
    )

    service.get_details(20, user=_ADMIN)
    assert _contract_owner(db_path, "0882/09/25") is None

    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "UPDATE kp_meta SET status = ?, paid_at = ? WHERE kp_id = 20",
            ("на согласовании", "2026-09-28T10:00:00"),
        )
        conn.execute(
            """
            UPDATE KP_offers
            SET order_number_1c = ?, specification_json = ?
            WHERE kp_id = 20
            """,
            ("ЯР-1", _SPEC_JSON),
        )
        conn.commit()

    details = service.move_to_production(20, "15.10.2026", user=_ADMIN)

    assert details.status == "в работе"
    assert _contract_owner(db_path, "0882/09/25") is None
