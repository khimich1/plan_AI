"""Commercial pile/FBS preview must open PRICE_DB_PATH, not PROJECT_ROOT/pb.db."""

from __future__ import annotations

import os
import sqlite3
import subprocess
import sys
from pathlib import Path

import pandas as pd
import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]
DISTINCTIVE_PILE_PRICE = 99999.12


def _write_sample_pile_xlsx(path: Path) -> None:
    rows = [
        [None, "Наименование", 15, 20, 22.5, 25, "30 на граните"],
        [None, "35 СЕЧЕНИЕ", None, None, None, None, None],
        [69, "С120.35-12", 11111.11, 22222.22, 33333.33, DISTINCTIVE_PILE_PRICE, 55555.55],
    ]
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Прайс", index=False, header=False)


def _seed_pile_db(tmp_path: Path) -> Path:
    xlsx_path = tmp_path / "piles.xlsx"
    db_path = tmp_path / "data" / "pb.db"
    db_path.parent.mkdir(parents=True, exist_ok=True)
    _write_sample_pile_xlsx(xlsx_path)
    from core.pile_price_db import import_pile_prices_from_xlsx

    import_pile_prices_from_xlsx(str(xlsx_path), str(db_path))
    return db_path


def _align_db_path_from_price_db_path(
    monkeypatch: pytest.MonkeyPatch,
    pb_db: Path,
) -> str:
    """Re-bind import-time snapshots the same way production does: DB_PATH = str(PRICE_DB_PATH)."""
    import app.services.commercial_calculation_service as commercial_calculation_service
    import app.services.commercial_workflow_service as commercial_workflow_service
    import app.services.product_draft_config as product_draft_config
    import core.commercial_offer as commercial_offer
    import core.commercial_offer_xlsx as commercial_offer_xlsx
    import core.price_db as price_db
    import core.project_paths as project_paths

    monkeypatch.setenv("PB_DB_PATH", str(pb_db))
    monkeypatch.setenv("PRICE_DB_PATH", str(pb_db))
    monkeypatch.setattr(project_paths, "PRICE_DB_PATH", Path(pb_db))
    db_str = str(project_paths.PRICE_DB_PATH)
    monkeypatch.setattr(commercial_offer_xlsx, "DB_PATH", db_str)
    monkeypatch.setattr(commercial_offer, "DB_PATH", db_str)
    monkeypatch.setattr(price_db, "DEFAULT_DB", db_str)
    monkeypatch.setattr(product_draft_config, "DB_PATH", db_str)
    monkeypatch.setattr(commercial_workflow_service, "DB_PATH", db_str)
    monkeypatch.setattr(commercial_calculation_service, "DB_PATH", db_str)
    return db_str


def _guard_project_root_pb_connect(monkeypatch: pytest.MonkeyPatch) -> None:
    """Simulate missing /app/pb.db: sqlite must not open PROJECT_ROOT/pb.db."""
    import core.project_paths as project_paths

    root_pb = (project_paths.PROJECT_ROOT / "pb.db").resolve()
    real_connect = sqlite3.connect

    def guarded(database, *args, **kwargs):  # type: ignore[no-untyped-def]
        if database not in (":memory:", None):
            try:
                if Path(str(database)).resolve() == root_pb:
                    raise sqlite3.OperationalError("unable to open database file")
            except (OSError, ValueError):
                pass
        return real_connect(database, *args, **kwargs)

    monkeypatch.setattr(sqlite3, "connect", guarded)


def test_db_path_follows_pb_db_path_at_import(tmp_path: Path) -> None:
    alt = tmp_path / "data" / "pb.db"
    alt.parent.mkdir(parents=True, exist_ok=True)
    alt.write_bytes(b"")

    env = os.environ.copy()
    env.pop("PRICE_DB_PATH", None)
    env["PB_DB_PATH"] = str(alt)
    pythonpath = str(REPO_ROOT)
    existing = env.get("PYTHONPATH", "")
    env["PYTHONPATH"] = pythonpath if not existing else pythonpath + os.pathsep + existing

    script = r"""
from pathlib import Path
from core.project_paths import PRICE_DB_PATH, PROJECT_ROOT
from core.commercial_offer_xlsx import DB_PATH as XLSX_DB
from core.commercial_offer import DB_PATH as PDF_DB
from core.price_db import DEFAULT_DB

expected = Path(r'''%s''')
assert Path(PRICE_DB_PATH) == expected, PRICE_DB_PATH
assert Path(XLSX_DB) == expected, XLSX_DB
assert Path(PDF_DB) == expected, PDF_DB
assert Path(DEFAULT_DB) == expected, DEFAULT_DB
assert Path(XLSX_DB) != PROJECT_ROOT / "pb.db"
print("ok")
""" % alt

    result = subprocess.run(
        [sys.executable, "-c", script],
        cwd=str(REPO_ROOT),
        env=env,
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode == 0, result.stderr + result.stdout
    assert "ok" in result.stdout


def test_get_pile_price_uses_env_db_not_project_root(tmp_path: Path) -> None:
    pb_db = _seed_pile_db(tmp_path)

    env = os.environ.copy()
    env.pop("PRICE_DB_PATH", None)
    env["PB_DB_PATH"] = str(pb_db)
    pythonpath = str(REPO_ROOT)
    existing = env.get("PYTHONPATH", "")
    env["PYTHONPATH"] = pythonpath if not existing else pythonpath + os.pathsep + existing

    script = r"""
from core.pile_price_db import get_pile_price
price = get_pile_price("С120.35-12", "B25")
assert price is not None, price
assert abs(float(price) - %s) < 0.01, price
print("ok")
""" % DISTINCTIVE_PILE_PRICE

    result = subprocess.run(
        [sys.executable, "-c", script],
        cwd=str(REPO_ROOT),
        env=env,
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode == 0, result.stderr + result.stdout
    assert "ok" in result.stdout


def test_pile_preview_uses_env_db_when_project_root_pb_missing(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    pb_db = _seed_pile_db(tmp_path)
    _align_db_path_from_price_db_path(monkeypatch, pb_db)
    _guard_project_root_pb_connect(monkeypatch)

    import core.commercial_offer_xlsx as commercial_offer_xlsx
    import core.price_db as price_db
    from app.services.commercial_pile_service import CommercialPileService
    from core.pile_price_db import get_pile_price

    assert Path(commercial_offer_xlsx.DB_PATH) == Path(pb_db)
    assert Path(price_db.DEFAULT_DB) == Path(pb_db)
    assert Path(commercial_offer_xlsx.DB_PATH) != Path(REPO_ROOT) / "pb.db"

    price = get_pile_price("С120.35-12", "B25", db_path=price_db.DEFAULT_DB)
    assert price == pytest.approx(DISTINCTIVE_PILE_PRICE)

    preview = CommercialPileService().generate_preview(
        "С120.35-12 2",
        db_path=str(commercial_offer_xlsx.DB_PATH),
    )
    assert preview.order_data, preview
    assert preview.order_data[0]["unit_price"] == pytest.approx(DISTINCTIVE_PILE_PRICE)


@pytest.mark.parametrize("product", ["piles", "fbs"])
def test_sibling_preview_does_not_connect_project_root_pb(
    product: str,
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    pb_db = tmp_path / "data" / "pb.db"
    pb_db.parent.mkdir(parents=True, exist_ok=True)
    sqlite3.connect(str(pb_db)).close()

    _align_db_path_from_price_db_path(monkeypatch, pb_db)
    _guard_project_root_pb_connect(monkeypatch)

    import core.commercial_offer_xlsx as commercial_offer_xlsx

    db_path = str(commercial_offer_xlsx.DB_PATH)
    if product == "piles":
        from app.services.commercial_pile_service import CommercialPileService

        preview = CommercialPileService().generate_preview("С120.35-12 1", db_path=db_path)
    else:
        from app.services.commercial_fbs_service import CommercialFbsService

        preview = CommercialFbsService().generate_preview("ФБС 9.3.6-Т 1", db_path=db_path)

    assert preview.order_data
    assert preview.order_data[0]["unit_price"] is None


def test_preview_piles_create_draft_path_uses_price_db_path(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    """create_draft → _preview_piles(db_path=str(DB_PATH)); no FastAPI app import."""
    pb_db = _seed_pile_db(tmp_path)
    _align_db_path_from_price_db_path(monkeypatch, pb_db)
    _guard_project_root_pb_connect(monkeypatch)

    from app.services.commercial_pile_service import CommercialPileService
    from app.services.product_draft_config import _preview_piles

    class _Wf:
        pile_service = CommercialPileService()

    preview = _preview_piles(_Wf(), "С120.35-12 2")
    assert preview.order_data
    assert preview.order_data[0]["unit_price"] == pytest.approx(DISTINCTIVE_PILE_PRICE)
