"""CTR-002: CounterpartiesRepository — search, get, upsert."""

from __future__ import annotations

from pathlib import Path

from app.repositories.counterparties_repository import CounterpartiesRepository
from tests.helpers import kp_db_fixtures as fx


def _repo(tmp_path: Path) -> CounterpartiesRepository:
    return CounterpartiesRepository(db_path=fx.make_iso_db(tmp_path))


def _seed(repo: CounterpartiesRepository, rows: list[dict]) -> None:
    repo.upsert(rows)


def test_search_returns_only_active_clients_by_default(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    _seed(
        repo,
        [
            {
                "code_1c": "00-1",
                "name": "СТОЛИЦА ООО",
                "inn": "7701000001",
                "kpp": "770101001",
                "is_client": True,
            },
            {
                "code_1c": "00-2",
                "name": "СТОЛИЦА ПОСТАВЩИК",
                "inn": "7701000002",
                "is_client": False,
            },
            {
                "code_1c": "00-3",
                "name": "СТОЛИЦА АРХИВ",
                "inn": "7701000003",
                "is_client": True,
                "is_active": False,
            },
        ],
    )

    items = repo.search("сто")
    assert [row["code_1c"] for row in items] == ["00-1"]

    all_items = repo.search("сто", clients_only=False)
    codes = {row["code_1c"] for row in all_items}
    assert codes == {"00-1", "00-2"}
    assert "00-3" not in codes


def test_search_ranks_prefix_then_substring_then_inn(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    _seed(
        repo,
        [
            {
                "code_1c": "00-sub",
                "name": "ТРЕСТ СТО ООО",
                "inn": "1111111111",
                "is_client": True,
            },
            {
                "code_1c": "00-pref",
                "name": "СТОЛИЦА ООО",
                "inn": "2222222222",
                "is_client": True,
            },
        ],
    )
    letter_items = repo.search("сто")
    assert [row["code_1c"] for row in letter_items] == ["00-pref", "00-sub"]

    _seed(
        repo,
        [
            {
                "code_1c": "12-sub",
                "name": "ТРЕСТ 12 ООО",
                "inn": "9900000000",
                "is_client": True,
            },
            {
                "code_1c": "12-pref",
                "name": "12 МЕСЯЦЕВ ООО",
                "inn": "8800000000",
                "is_client": True,
            },
            {
                "code_1c": "12-inn",
                "name": "РОМАШКА ООО",
                "inn": "1234567890",
                "is_client": True,
            },
        ],
    )
    digit_items = repo.search("12")
    assert [row["code_1c"] for row in digit_items] == ["12-pref", "12-sub", "12-inn"]


def test_search_is_casefold_and_trims_query(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    _seed(
        repo,
        [
            {
                "code_1c": "00-1",
                "name": "  СТОЛИЦА ООО  ",
                "inn": "7701000001",
                "is_client": True,
            }
        ],
    )
    stored = repo.get_by_code_1c("00-1")
    assert stored is not None
    assert stored["name"] == "СТОЛИЦА ООО"
    assert stored["name_normalized"] == "столица ооо"

    items = repo.search("  СтО  ")
    assert len(items) == 1
    assert items[0]["code_1c"] == "00-1"


def test_search_inn_prefix_when_query_is_digits(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    _seed(
        repo,
        [
            {
                "code_1c": "00-a",
                "name": "АЛЬФА",
                "inn": "7707123456",
                "is_client": True,
            },
            {
                "code_1c": "00-b",
                "name": "БЕТА",
                "inn": "9900000000",
                "is_client": True,
            },
        ],
    )
    items = repo.search("7707")
    assert [row["code_1c"] for row in items] == ["00-a"]


def test_search_respects_limit(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    _seed(
        repo,
        [
            {
                "code_1c": f"00-{i:02d}",
                "name": f"СТО Клиент {i}",
                "is_client": True,
            }
            for i in range(12)
        ],
    )
    items = repo.search("сто", limit=10)
    assert len(items) == 10


def test_get_by_id_and_code(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    stats = repo.upsert(
        [{"code_1c": "Б00000002", "name": "ИП Иванов", "inn": "123456789012"}]
    )
    assert stats.added == 1
    row = repo.get_by_code_1c("Б00000002")
    assert row is not None
    by_id = repo.get_by_id(int(row["id"]))
    assert by_id is not None
    assert by_id["name"] == "ИП Иванов"
    assert repo.get_by_id(999999) is None
    assert repo.get_by_code_1c("missing") is None


def test_upsert_inserts_updates_and_skips_unchanged(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    first = repo.upsert(
        [
            {
                "code_1c": "00-1",
                "name": "РОМАШКА",
                "inn": "7701000001",
                "kpp": "770101001",
                "is_client": True,
            }
        ]
    )
    assert first.added == 1
    assert first.updated == 0
    assert first.unchanged == 0
    original = repo.get_by_code_1c("00-1")
    assert original is not None
    original_updated_at = original["updated_at"]

    same = repo.upsert(
        [
            {
                "code_1c": "00-1",
                "name": "РОМАШКА",
                "inn": "7701000001",
                "kpp": "770101001",
                "is_client": True,
            }
        ]
    )
    assert same.added == 0
    assert same.updated == 0
    assert same.unchanged == 1
    after_same = repo.get_by_code_1c("00-1")
    assert after_same is not None
    assert after_same["updated_at"] == original_updated_at

    changed = repo.upsert(
        [
            {
                "code_1c": "00-1",
                "name": "РОМАШКА НОВАЯ",
                "inn": "7701000001",
                "kpp": "770101001",
                "is_client": True,
            }
        ]
    )
    assert changed.added == 0
    assert changed.updated == 1
    assert changed.unchanged == 0
    updated = repo.get_by_code_1c("00-1")
    assert updated is not None
    assert updated["name"] == "РОМАШКА НОВАЯ"
    assert updated["name_normalized"] == "ромашка новая"


def test_upsert_stores_empty_inn_kpp_as_null(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    repo.upsert([{"code_1c": "00-empty", "name": "ИП Без ИНН", "inn": "", "kpp": "  "}])
    row = repo.get_by_code_1c("00-empty")
    assert row is not None
    assert row["inn"] is None
    assert row["kpp"] is None
