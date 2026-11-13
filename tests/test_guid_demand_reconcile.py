"""NG-001…NG-004: сверка очереди guid_demand с КП (две временные базы)."""

from __future__ import annotations

import hashlib
import json
import sqlite3
import subprocess
import sys
from pathlib import Path

import pytest

from core.guid_demand import (
    FIELD_GUID,
    FIELD_GUID_U,
    STATE_OPEN,
    STATE_RESOLVED,
    ensure_schema as ensure_demand_schema,
)
from core.guid_demand_reconcile import (
    ACTION_CLOSED,
    ACTION_NARROWED,
    ACTION_OPENED,
    ACTION_REBOUND,
    format_report,
    reconcile_guid_demand,
)
from core.nomenclature_guid import ensure_schema as ensure_guid_schema
from core.nomenclature_guid import upsert
from core.plate_guid_choice import ensure_schema as ensure_choice_schema
from tests.helpers import kp_db_fixtures as fx

PROJECT_ROOT = Path(__file__).resolve().parents[1]
READY_GUID = "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee"
PLATE_GUID = "11111111-1111-1111-1111-111111111111"

_NON_PLATE_TABLES = {
    "pile": "kp_piles",
    "stair_step": "kp_steps",
    "stair_flight": "kp_marches",
    "bridge_pile": "kp_bridge_piles",
    "composite_pile": "kp_composite_piles",
    "fbs": "kp_fbs",
}


@pytest.fixture
def plita(tmp_path: Path) -> str:
    return fx.make_iso_db(tmp_path)


@pytest.fixture
def pb(tmp_path: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(tmp_path / "pb.db")
    ensure_guid_schema(conn)
    ensure_choice_schema(conn)
    ensure_demand_schema(conn)
    conn.execute(
        'CREATE TABLE IF NOT EXISTS prays_plity ('
        '"Уникальный идентификатор (Номенклатура)" TEXT, "Товар" TEXT)'
    )
    conn.commit()
    try:
        yield conn
    finally:
        conn.close()


def _kp(plita: str, kp_id: int, status: str = "в архиве") -> None:
    fx.seed_kp_offer(plita, kp_id, status=status)


def _plate(plita: str, kp_id: int, name: str) -> None:
    fx.seed_plate(
        plita,
        kp_id=kp_id,
        plate_name=name,
        length_m=6.0,
        width_m=1.2,
        qty=1,
        status="в производстве",
    )


def _line(plita: str, kp_id: int, kind: str, mark: str) -> None:
    if kind == "plate":
        _plate(plita, kp_id, mark)
        return
    table = _NON_PLATE_TABLES[kind]
    columns = "kp_id, position_number, mark, qty, unit_price, discounted_price"
    values = "?, 1, ?, 1, 100, 100"
    if kind not in {"stair_step"}:
        columns += ", concrete_grade"
        values += ", 'B25'"
    with sqlite3.connect(plita) as conn:
        conn.execute(f"INSERT INTO {table} ({columns}) VALUES ({values})", (kp_id, mark))
        conn.commit()


def _demand(
    pb: sqlite3.Connection,
    kind: str,
    mark: str,
    kp_ids: list[int],
    *,
    field: str = FIELD_GUID,
    state: str = STATE_OPEN,
) -> None:
    pb.execute(
        "INSERT INTO guid_demand (product_kind, mark, field, reason, hint, state, "
        "kp_ids_json, opened_at, resolved_at) VALUES (?, ?, ?, 'нет GUID 1С', "
        "'заведите карточку в 1С и загрузите отчёт', ?, json(?), '2026-11-01T00:00:00Z', ?)",
        (
            kind,
            mark,
            field,
            state,
            json.dumps(kp_ids),
            None if state == STATE_OPEN else "2026-11-02T00:00:00Z",
        ),
    )
    pb.commit()


def _ready_pile(pb: sqlite3.Connection, mark: str) -> None:
    upsert(pb, "pile", mark, guid_1c=READY_GUID, match_status="auto")
    pb.commit()


def _rows(pb: sqlite3.Connection) -> dict[tuple[str, str, str], tuple[str, list[int]]]:
    out = {}
    for kind, mark, field, state, ids in pb.execute(
        "SELECT product_kind, mark, field, state, kp_ids_json FROM guid_demand"
    ):
        out[(kind, mark, field)] = (state, json.loads(ids))
    return out


def _reconcile(pb: sqlite3.Connection, plita: str):
    plita_conn = sqlite3.connect(f"file:{plita}?mode=ro", uri=True)
    try:
        return reconcile_guid_demand(pb, plita_conn)
    finally:
        plita_conn.close()


def test_narrows_kp_list_to_kps_that_still_hold_the_mark(plita, pb) -> None:
    mark = "Плиты ПБ 64,40-10,2-1п"
    for kp_id in (55, 57, 59):
        _kp(plita, kp_id)
    _plate(plita, 55, mark)
    _plate(plita, 57, mark)
    _plate(plita, 59, "Плиты ПБ 10-10-8п")
    pb.execute(
        'INSERT INTO prays_plity VALUES (?, ?)', (PLATE_GUID, "Плиты ПБ 10-10-8п")
    )
    _demand(pb, "plate", mark, [55, 57, 59])

    report = _reconcile(pb, plita)

    assert _rows(pb)[("plate", mark, FIELD_GUID)] == (STATE_OPEN, [55, 57])
    change = report.changes[0]
    assert (change.action, change.kp_ids_before, change.kp_ids_after) == (
        ACTION_NARROWED,
        (55, 57, 59),
        (55, 57),
    )
    assert len(report.changes) == 1  # КП 59 перепроверен, у плиты есть GUID


def test_closes_row_when_kp_is_deleted(plita, pb) -> None:
    _demand(pb, "pile", "С99.35-9", [138])

    report = _reconcile(pb, plita)

    state, ids = _rows(pb)[("pile", "С99.35-9", FIELD_GUID)]
    assert state == STATE_RESOLVED and ids == []
    assert [c.action for c in report.changes] == [ACTION_CLOSED]
    resolved_at = pb.execute("SELECT resolved_at FROM guid_demand").fetchone()[0]
    assert resolved_at


def test_closes_row_after_mark_edit_and_opens_new_row_without_number(plita, pb) -> None:
    _kp(plita, 124)
    _plate(plita, 124, "Плиты ПБ 97-10,2-6п")
    _demand(pb, "plate", "Плиты ПБ 97-10,2-8п", [124])

    report = _reconcile(pb, plita)

    rows = _rows(pb)
    assert rows[("plate", "Плиты ПБ 97-10,2-8п", FIELD_GUID)][0] == STATE_RESOLVED
    assert rows[("plate", "Плиты ПБ 97-10,2-6п", FIELD_GUID)] == (STATE_OPEN, [124])
    assert [c.action for c in report.changes] == [ACTION_CLOSED, ACTION_OPENED]
    assert report.changes[1].mark == "Плиты ПБ 97-10,2-6п"
    assert report.changes[1].kp_ids_after == (124,)


def test_does_not_open_new_row_when_new_mark_already_has_number(plita, pb) -> None:
    _kp(plita, 126)
    _line(plita, 126, "stair_flight", "ЛС11-1")
    upsert(pb, "stair_flight", "ЛС11-1", guid_1c=READY_GUID, match_status="auto")
    pb.commit()
    _demand(pb, "stair_flight", "ЛС11", [126])

    report = _reconcile(pb, plita)

    rows = _rows(pb)
    assert rows[("stair_flight", "ЛС11", FIELD_GUID)][0] == STATE_RESOLVED
    assert ("stair_flight", "ЛС11-1", FIELD_GUID) not in rows
    assert [c.action for c in report.changes] == [ACTION_CLOSED]


def test_rebinds_row_to_kp_that_holds_same_mark(plita, pb) -> None:
    mark = "С120.35-12"
    _kp(plita, 9)
    _line(plita, 9, "pile", mark)
    _demand(pb, "pile", mark, [138])

    report = _reconcile(pb, plita)

    assert _rows(pb)[("pile", mark, FIELD_GUID)] == (STATE_OPEN, [9])
    change = report.changes[0]
    assert (change.action, change.kp_ids_before, change.kp_ids_after) == (
        ACTION_REBOUND,
        (138,),
        (9,),
    )


def test_second_run_changes_nothing(plita, pb) -> None:
    _kp(plita, 124)
    _plate(plita, 124, "Плиты ПБ 97-10,2-6п")
    _demand(pb, "plate", "Плиты ПБ 97-10,2-8п", [124])
    _demand(pb, "pile", "С99.35-9", [138])
    first = _reconcile(pb, plita)
    snapshot = _rows(pb)

    second = _reconcile(pb, plita)

    assert first.has_changes
    assert not second.has_changes
    assert second.changes == ()
    assert _rows(pb) == snapshot


def test_unchanged_rows_are_counted_and_not_touched(plita, pb) -> None:
    _kp(plita, 3)
    _line(plita, 3, "pile", "С70.35-9у")
    _demand(pb, "pile", "С70.35-9у", [3], field=FIELD_GUID_U)

    report = _reconcile(pb, plita)

    assert report.changes == ()
    assert report.unchanged == 1
    assert _rows(pb)[("pile", "С70.35-9у", FIELD_GUID_U)] == (STATE_OPEN, [3])


def test_u_mark_is_matched_exactly_with_suffix(plita, pb) -> None:
    _kp(plita, 3)
    _line(plita, 3, "pile", "С70.35-9")  # без «у»
    _ready_pile(pb, "С70.35-9")
    _demand(pb, "pile", "С70.35-9у", [3], field=FIELD_GUID_U)

    report = _reconcile(pb, plita)

    assert _rows(pb)[("pile", "С70.35-9у", FIELD_GUID_U)][0] == STATE_RESOLVED
    assert [c.action for c in report.changes] == [ACTION_CLOSED]


def test_latin_and_cyrillic_letters_are_different_marks(plita, pb) -> None:
    _kp(plita, 47)
    _line(plita, 47, "bridge_pile", "С14-40T4")  # кириллическая С
    _demand(pb, "bridge_pile", "C14-40T4", [47])  # латинская C

    report = _reconcile(pb, plita)

    rows = _rows(pb)
    assert rows[("bridge_pile", "C14-40T4", FIELD_GUID)][0] == STATE_RESOLVED
    assert rows[("bridge_pile", "С14-40T4", FIELD_GUID)] == (STATE_OPEN, [47])
    assert [c.action for c in report.changes] == [ACTION_CLOSED, ACTION_OPENED]


def test_reopens_resolved_row_when_dropped_kp_needs_it_again(plita, pb) -> None:
    _kp(plita, 7)
    _line(plita, 7, "fbs", "ФБС 24.4.6-Т")
    _demand(pb, "fbs", "ФБС 24.4.6-Т", [8], state=STATE_RESOLVED)
    _demand(pb, "pile", "С1.1-1", [7])

    _reconcile(pb, plita)

    state, ids = _rows(pb)[("fbs", "ФБС 24.4.6-Т", FIELD_GUID)]
    assert state == STATE_OPEN and ids == [7]  # хвост КП 8 не тянем
    opened_at = pb.execute(
        "SELECT opened_at FROM guid_demand WHERE product_kind = 'fbs'"
    ).fetchone()[0]
    assert opened_at == "2026-11-01T00:00:00Z"


@pytest.mark.parametrize("kind", sorted(_NON_PLATE_TABLES))
def test_every_kind_is_read_from_its_own_table(plita, pb, kind: str) -> None:
    _kp(plita, 5)
    _line(plita, 5, kind, "МАРКА-1")
    _demand(pb, kind, "МАРКА-1", [5, 6])

    _reconcile(pb, plita)

    assert _rows(pb)[(kind, "МАРКА-1", FIELD_GUID)] == (STATE_OPEN, [5])


def test_kp_in_any_status_counts(plita, pb) -> None:
    _kp(plita, 59, status="на согласовании")
    _line(plita, 59, "pile", "С1.1-1")
    _demand(pb, "pile", "С1.1-1", [59])

    report = _reconcile(pb, plita)

    assert report.changes == ()


def test_reconcile_does_not_commit(plita, pb, tmp_path: Path) -> None:
    _demand(pb, "pile", "С99.35-9", [138])

    _reconcile(pb, plita)
    pb.rollback()

    assert _rows(pb)[("pile", "С99.35-9", FIELD_GUID)] == (STATE_OPEN, [138])


def test_format_report_is_russian_and_lists_each_change(plita, pb) -> None:
    mark = "С120.35-12"
    _kp(plita, 9)
    _line(plita, 9, "pile", mark)
    _demand(pb, "pile", mark, [138])
    _demand(pb, "pile", "С99.35-9", [140])

    text = format_report(_reconcile(pb, plita), apply=False)

    assert "сухой прогон" in text.lower()
    assert "С120.35-12" in text and "138" in text and "9" in text
    assert "закрыта" in text and "перепривязана" in text
    assert "Не изменилось: 0" in text


def test_format_report_for_clean_run() -> None:
    from core.guid_demand_reconcile import ReconcileReport

    text = format_report(ReconcileReport(changes=(), unchanged=52), apply=True)

    assert "изменений нет" in text.lower()
    assert "52" in text


# ---------------------------------------------------------------- NG-004 script


def _file_hash(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _script(*args: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, str(PROJECT_ROOT / "scripts" / "reconcile_guid_demand.py"), *args],
        capture_output=True,
        text=True,
        cwd=PROJECT_ROOT,
        check=False,
    )


@pytest.fixture
def copies(plita, pb, tmp_path: Path) -> tuple[Path, Path]:
    mark = "С120.35-12"
    _kp(plita, 9)
    _line(plita, 9, "pile", mark)
    _demand(pb, "pile", mark, [138])
    _demand(pb, "pile", "С99.35-9", [140])
    pb.close()
    return tmp_path / "pb.db", Path(plita)


def test_script_default_is_dry_run_and_does_not_touch_files(copies) -> None:
    pb_path, plita_path = copies
    before = (_file_hash(pb_path), _file_hash(plita_path))

    result = _script("--pb-db", str(pb_path), "--plita-db", str(plita_path))

    assert result.returncode == 0, result.stderr
    assert (_file_hash(pb_path), _file_hash(plita_path)) == before
    assert "сухой прогон" in result.stdout.lower()
    assert str(pb_path) in result.stdout and str(plita_path) in result.stdout
    assert "С120.35-12" in result.stdout


def test_script_apply_changes_then_second_run_is_empty(copies) -> None:
    pb_path, plita_path = copies

    applied = _script("--pb-db", str(pb_path), "--plita-db", str(plita_path), "--apply")
    again = _script("--pb-db", str(pb_path), "--plita-db", str(plita_path), "--apply")

    assert applied.returncode == 0, applied.stderr
    assert again.returncode == 0, again.stderr
    assert "изменений нет" in again.stdout.lower()
    with sqlite3.connect(pb_path) as conn:
        states = dict(conn.execute("SELECT mark, state FROM guid_demand"))
    assert states == {"С120.35-12": STATE_OPEN, "С99.35-9": STATE_RESOLVED}


def test_script_writes_report_file(copies, tmp_path: Path) -> None:
    pb_path, plita_path = copies
    report_path = tmp_path / "report.txt"

    result = _script(
        "--pb-db", str(pb_path), "--plita-db", str(plita_path), "--report", str(report_path)
    )

    assert result.returncode == 0, result.stderr
    assert "С99.35-9" in report_path.read_text(encoding="utf-8")


def test_script_help_warns_about_copies() -> None:
    result = _script("--help")

    assert result.returncode == 0
    assert "копи" in result.stdout.lower()


def test_script_fails_on_missing_database(tmp_path: Path) -> None:
    result = _script(
        "--pb-db", str(tmp_path / "нет.db"), "--plita-db", str(tmp_path / "нет2.db")
    )

    assert result.returncode != 0


# ------------------------------------------------- NG-002: обёртка и вызовы

import logging  # noqa: E402

from app.services import guid_queue_sync  # noqa: E402
from app.services.guid_queue_sync import sync_guid_queue_with_kp  # noqa: E402


def _closed_pair(plita, pb) -> Path:
    """КП 138 нет в plita.db, строка очереди про него открыта."""
    _demand(pb, "pile", "С99.35-9", [138])
    pb.close()
    return Path(plita).parent / "pb.db"


def test_sync_closes_stale_row_and_commits(plita, pb) -> None:
    pb_path = _closed_pair(plita, pb)

    report = sync_guid_queue_with_kp(pb_path, plita)

    assert report is not None and report.has_changes
    with sqlite3.connect(pb_path) as conn:
        assert conn.execute("SELECT state FROM guid_demand").fetchone() == (STATE_RESOLVED,)


def test_sync_without_changes_does_not_write_pb(plita, pb) -> None:
    _kp(plita, 3)
    _line(plita, 3, "pile", "С70.35-9")
    _demand(pb, "pile", "С70.35-9", [3])
    pb.close()
    pb_path = Path(plita).parent / "pb.db"
    before = _file_hash(pb_path)

    report = sync_guid_queue_with_kp(pb_path, plita)

    assert report is not None and not report.has_changes
    assert _file_hash(pb_path) == before


@pytest.mark.parametrize("which", ["pb", "plita"])
def test_sync_returns_none_when_database_missing(plita, pb, which: str, tmp_path) -> None:
    pb.close()
    pb_path = tmp_path / "pb.db"
    plita_path = Path(plita)
    if which == "pb":
        pb_path = tmp_path / "нет-pb.db"
    else:
        plita_path = tmp_path / "нет-plita.db"

    assert sync_guid_queue_with_kp(pb_path, plita_path) is None


def test_sync_never_raises_and_logs_when_reconcile_fails(
    plita, pb, monkeypatch: pytest.MonkeyPatch, caplog: pytest.LogCaptureFixture
) -> None:
    pb_path = _closed_pair(plita, pb)

    def _boom(*_args, **_kwargs):
        raise RuntimeError("сбой сверки")

    monkeypatch.setattr(guid_queue_sync, "reconcile_guid_demand", _boom)
    with caplog.at_level(logging.ERROR):
        assert sync_guid_queue_with_kp(pb_path, plita) is None

    assert "сверить очередь" in caplog.text
    with sqlite3.connect(pb_path) as conn:
        assert conn.execute("SELECT state FROM guid_demand").fetchone() == (STATE_OPEN,)


def test_sync_survives_locked_pb(
    plita, pb, monkeypatch: pytest.MonkeyPatch, caplog: pytest.LogCaptureFixture
) -> None:
    pb_path = _closed_pair(plita, pb)
    monkeypatch.setattr(guid_queue_sync, "_PB_TIMEOUT_SEC", 0.05)
    locker = sqlite3.connect(pb_path)
    locker.execute("BEGIN EXCLUSIVE")
    try:
        with caplog.at_level(logging.ERROR):
            assert sync_guid_queue_with_kp(pb_path, plita) is None
    finally:
        locker.rollback()
        locker.close()

    assert "сверить очередь" in caplog.text


@pytest.fixture
def archive_env(plita, pb, tmp_path, monkeypatch: pytest.MonkeyPatch) -> Path:
    from app.core.settings import get_settings

    pb_path = tmp_path / "pb.db"
    monkeypatch.setenv("PB_DB_PATH", str(pb_path))
    monkeypatch.setenv("PLITA_DB_PATH", str(plita))
    get_settings.cache_clear()
    pb.close()
    return pb_path


def test_delete_offer_closes_queue_row_of_deleted_kp(plita, archive_env, tmp_path) -> None:
    from unittest.mock import MagicMock

    from app.repositories.kp_archive_repository import KpArchiveRepository
    from app.services.archive_service import ArchiveService

    _kp(plita, 138)
    _line(plita, 138, "pile", "С99.35-9")
    with sqlite3.connect(archive_env) as conn:
        _demand(conn, "pile", "С99.35-9", [138])
    service = ArchiveService(
        repository=KpArchiveRepository(db_path=str(plita)),
        outputs_dir=tmp_path,
        promise_service=MagicMock(),
    )

    service.delete_offer(138, user={"id": 1, "role": "admin"})

    with sqlite3.connect(archive_env) as conn:
        assert conn.execute("SELECT state FROM guid_demand").fetchone() == (STATE_RESOLVED,)


def test_delete_offer_failure_does_not_sync(
    tmp_path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from unittest.mock import MagicMock

    from app.services import archive_service
    from app.services.archive_service import ArchiveNotFoundError, ArchiveService

    calls: list[tuple] = []
    monkeypatch.setattr(
        archive_service, "sync_guid_queue_with_kp", lambda *a: calls.append(a)
    )
    repository = MagicMock()
    repository.db_path = str(tmp_path / "plita.db")
    repository.get_by_id.return_value = {"kp_id": 42, "owner_user_id": 1, "status": "в архиве"}
    repository.delete.return_value = False
    service = ArchiveService(
        repository=repository, outputs_dir=tmp_path, promise_service=MagicMock()
    )

    with pytest.raises(ArchiveNotFoundError):
        service.delete_offer(42, user={"id": 1, "role": "admin"})

    assert calls == []


def test_offers_service_create_syncs_queue_after_notify(
    tmp_path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from unittest.mock import MagicMock

    from app.services import kp_guid_notify, offers_service

    order: list[str] = []
    monkeypatch.setattr(
        kp_guid_notify, "notify_kp_guid_missing", lambda **_k: order.append("notify")
    )
    monkeypatch.setattr(
        offers_service, "sync_guid_queue_with_kp", lambda *a: order.append("sync")
    )
    repository = MagicMock()
    repository.db_path = str(tmp_path / "plita.db")
    service = offers_service.OffersService(
        kp_repository=repository, counterparties=MagicMock()
    )

    service._notify_guid_missing(7, [])

    assert order == ["notify", "sync"]


def test_offers_service_sync_runs_even_if_notify_fails(
    tmp_path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from unittest.mock import MagicMock

    from app.services import kp_guid_notify, offers_service

    calls: list[tuple] = []

    def _fail(**_kwargs):
        raise RuntimeError("уведомление упало")

    monkeypatch.setattr(kp_guid_notify, "notify_kp_guid_missing", _fail)
    monkeypatch.setattr(
        offers_service, "sync_guid_queue_with_kp", lambda *a: calls.append(a)
    )
    repository = MagicMock()
    repository.db_path = str(tmp_path / "plita.db")
    service = offers_service.OffersService(
        kp_repository=repository, counterparties=MagicMock()
    )

    service._notify_guid_missing(7, [])

    assert len(calls) == 1


# --------------------------------------------- NG-003: «Завести в 1С» (list_tasks)


def test_list_tasks_reconciles_before_answering(plita, archive_env) -> None:
    from app.services.nomenclature_queue_service import NomenclatureQueueService

    _kp(plita, 9)
    _line(plita, 9, "pile", "С120.35-12")
    with sqlite3.connect(archive_env) as conn:
        _demand(conn, "pile", "С120.35-12", [138])
        _demand(conn, "pile", "С99.35-9", [140])

    tasks = NomenclatureQueueService().list_tasks()

    assert [(t.mark, t.kp_ids) for t in tasks.to_create_1c] == [("С120.35-12", [9])]


def test_list_tasks_repeated_calls_do_not_write(plita, archive_env) -> None:
    from app.services.nomenclature_queue_service import NomenclatureQueueService

    _kp(plita, 9)
    _line(plita, 9, "pile", "С120.35-12")
    with sqlite3.connect(archive_env) as conn:
        _demand(conn, "pile", "С120.35-12", [138])
    service = NomenclatureQueueService()
    service.list_tasks()
    after_first = _file_hash(archive_env)

    service.list_tasks()
    service.list_tasks()

    assert _file_hash(archive_env) == after_first


def test_list_tasks_answers_even_if_sync_fails(
    plita, archive_env, monkeypatch: pytest.MonkeyPatch
) -> None:
    from app.services.nomenclature_queue_service import NomenclatureQueueService

    with sqlite3.connect(archive_env) as conn:
        _demand(conn, "pile", "С99.35-9", [140])
    monkeypatch.setattr(
        guid_queue_sync,
        "reconcile_guid_demand",
        lambda *_a: (_ for _ in ()).throw(RuntimeError("сбой")),
    )

    tasks = NomenclatureQueueService().list_tasks()

    assert [t.mark for t in tasks.to_create_1c] == ["С99.35-9"]


def test_list_tasks_with_explicit_pb_only_does_not_sync(plita, archive_env, tmp_path) -> None:
    """Сервис с явной pb.db (тесты, скрипты) не лезет в чужую plita.db из настроек."""
    from app.services.nomenclature_queue_service import NomenclatureQueueService

    other_pb = tmp_path / "other-pb.db"
    other_pb.write_bytes(archive_env.read_bytes())
    with sqlite3.connect(other_pb) as conn:
        _demand(conn, "pile", "С99.35-9", [140])

    tasks = NomenclatureQueueService(other_pb).list_tasks()

    assert [t.mark for t in tasks.to_create_1c] == ["С99.35-9"]
