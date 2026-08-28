"""GSM waybill blank export: fill .xlsx → soffice → .xls → zip.

Decision D3 quirks (Phase 0):
  - Bare ``--convert-to xls`` (NOT 'xls:"MS Excel 97"') — impl_store 0x81a crash
  - Isolated profile ``-env:UserInstallation=file://...``
  - Timeout on every soffice call
  - Never modify originals under ``ГСМ/**`` — always copy template to temp
"""

from __future__ import annotations

import io
import json
import logging
import shutil
import subprocess
import tempfile
import zipfile
from datetime import date
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from app.repositories.gsm_repository import GsmRepository
from core.gsm.blank import (
    BlankDriver,
    BlankWaybill,
    fill_workbook,
    legs_from_route_items,
    vehicle_mark_label,
    waybill_export_filename,
)
from core.gsm.season import norm_for, parse_season_switches

logger = logging.getLogger(__name__)

PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_TEMPLATE = PROJECT_ROOT / "core" / "gsm" / "templates" / "waybill_blank.xlsx"
DEFAULT_SOFFICE_TIMEOUT_SEC = 120


class GsmExportError(Exception):
    """Export failure with stable machine code for HTTP mapping."""

    def __init__(self, message: str, *, code: str, details: dict[str, Any] | None = None) -> None:
        super().__init__(message)
        self.code = code
        self.details = details or {}


def run_soffice(args: list[str], workdir: Path, timeout: int) -> None:
    """LibreOffice headless with isolated user profile (Phase 0 pattern)."""
    profile = workdir / "lo_profile"
    cmd = [
        "soffice",
        "--headless",
        "--norestore",
        f"-env:UserInstallation=file://{profile}",
        *args,
    ]
    try:
        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=timeout,
            cwd=workdir,
        )
    except subprocess.TimeoutExpired as exc:
        raise GsmExportError(
            f"LibreOffice timed out after {timeout}s while exporting waybills",
            code="gsm_export_soffice_timeout",
            details={"timeout_sec": timeout},
        ) from exc
    except FileNotFoundError as exc:
        raise GsmExportError(
            "LibreOffice (soffice) is not installed or not on PATH",
            code="gsm_export_soffice_missing",
        ) from exc
    if proc.returncode != 0:
        stderr = (proc.stderr or "").strip()
        raise GsmExportError(
            f"soffice failed during waybill export: {stderr or proc.stdout or 'unknown error'}",
            code="gsm_export_soffice_failed",
            details={"returncode": proc.returncode, "stderr": stderr[:2000]},
        )


def convert_with_soffice(
    src: Path,
    fmt: str,
    outdir: Path,
    workdir: Path,
    timeout: int,
) -> Path:
    """``soffice --convert-to`` with result check.

    NB: explicit filter ``xls:"MS Excel 97"`` crashes (impl_store 0x81a);
    bare ``xls`` uses the same filter and works.
    """
    run_soffice(
        ["--convert-to", fmt, "--outdir", str(outdir), str(src)],
        workdir,
        timeout,
    )
    out = outdir / f"{src.stem}.{fmt}"
    if not out.exists():
        raise GsmExportError(
            f"soffice did not produce {out.name}",
            code="gsm_export_soffice_failed",
            details={"expected": str(out)},
        )
    return out


class GsmExportService:
    def __init__(
        self,
        repo: GsmRepository,
        *,
        template_path: Path | None = None,
        soffice_timeout_sec: int = DEFAULT_SOFFICE_TIMEOUT_SEC,
    ) -> None:
        self._repo = repo
        self._template_path = Path(template_path) if template_path else DEFAULT_TEMPLATE
        self._timeout = int(soffice_timeout_sec)

    def export_zip(
        self,
        *,
        vehicle_ids: list[int],
        period_from: date,
        period_to: date,
    ) -> tuple[bytes, str]:
        """Build zip of «ПЛ DD.MM.YY.xls»; return (bytes, download_filename)."""
        if period_to < period_from:
            raise GsmExportError(
                "period_to must be >= period_from",
                code="gsm_invalid_period",
            )
        if not vehicle_ids:
            raise GsmExportError(
                "vehicle_ids must be non-empty",
                code="gsm_invalid_period",
            )

        ids = sorted({int(v) for v in vehicle_ids})
        for vid in ids:
            if self._repo.get_vehicle(vid) is None:
                raise GsmExportError(
                    f"vehicle #{vid} not found",
                    code="gsm_vehicle_not_found",
                )

        waybills: list[dict[str, Any]] = []
        for vid in ids:
            waybills.extend(
                self._repo.list_waybills(
                    vehicle_id=vid,
                    period_from=period_from,
                    period_to=period_to,
                )
            )
        if not waybills:
            raise GsmExportError(
                "no waybills in period for selected vehicles",
                code="gsm_export_empty",
            )

        waybills.sort(key=lambda r: (str(r["date"]), int(r["vehicle_id"]), int(r["id"])))

        if not self._template_path.exists():
            raise GsmExportError(
                f"blank template not found: {self._template_path}",
                code="gsm_export_template_missing",
            )

        stations = {
            int(s["id"]): str(s["address"])
            for s in self._repo.list_stations()
        }
        switches_raw = self._repo.get_setting("season_switches")
        try:
            season_switches = parse_season_switches(switches_raw)
        except ValueError as exc:
            raise GsmExportError(
                f"invalid season_switches setting: {switches_raw!r}",
                code="gsm_settings_invalid",
            ) from exc

        buf = io.BytesIO()
        with tempfile.TemporaryDirectory(prefix="gsm_export_") as tmp:
            workdir = Path(tmp)
            outdir = workdir / "out"
            outdir.mkdir()
            filled_dir = workdir / "filled"
            filled_dir.mkdir()

            with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                used_names: dict[str, int] = {}
                for row in waybills:
                    day = _as_date(row["date"])
                    vehicle = self._repo.get_vehicle(int(row["vehicle_id"]))
                    assert vehicle is not None
                    driver = self._repo.get_driver(int(row["driver_id"]))
                    if driver is None:
                        raise GsmExportError(
                            f"driver #{row['driver_id']} not found",
                            code="gsm_driver_not_found",
                        )

                    norm = norm_for(
                        day,
                        norm_summer=float(vehicle["norm_summer"]),
                        norm_winter=float(vehicle["norm_winter"]),
                        switches=season_switches,
                    )
                    route_items = _parse_route_list(row.get("route_json"))
                    legs = legs_from_route_items(route_items, stations_by_id=stations)

                    blank = BlankWaybill(
                        day=day,
                        vehicle_mark=vehicle_mark_label(str(vehicle["name"])),
                        plate_number=str(vehicle["plate_number"]),
                        driver=BlankDriver(
                            full_name=str(driver["full_name"]),
                            license_number=str(driver["license_number"]),
                            license_issued_at=driver.get("license_issued_at"),
                            personnel_number=driver.get("personnel_number"),
                            snils=driver.get("snils"),
                        ),
                        odometer_start=int(row["odometer_start"] or 0),
                        fuel_start=float(row["fuel_start"] or 0.0),
                        fuel_issued=float(row["fuel_issued"] or 0.0),
                        norm_l_per_100=float(norm),
                        legs=legs,
                    )

                    # Copy template — never touch originals / shared template in-place.
                    xlsx_path = filled_dir / f"wb_{row['id']}_{day.isoformat()}.xlsx"
                    shutil.copy2(self._template_path, xlsx_path)
                    wb = load_workbook(xlsx_path, data_only=False)
                    fill_workbook(wb, blank)
                    wb.save(xlsx_path)

                    try:
                        xls_path = convert_with_soffice(
                            xlsx_path, "xls", outdir, workdir, self._timeout
                        )
                    except GsmExportError:
                        raise
                    except RuntimeError as exc:
                        raise GsmExportError(
                            str(exc),
                            code="gsm_export_soffice_failed",
                        ) from exc

                    entry_name = waybill_export_filename(day)
                    if entry_name in used_names:
                        used_names[entry_name] += 1
                        stem = entry_name[: -len(".xls")]
                        entry_name = f"{stem}_{used_names[entry_name]}.xls"
                    else:
                        used_names[entry_name] = 1

                    zf.writestr(entry_name, xls_path.read_bytes())

        # Status flip only after the full zip is built successfully.
        for row in waybills:
            day = _as_date(row["date"])
            self._repo.upsert_waybill(
                vehicle_id=int(row["vehicle_id"]),
                date=day,
                driver_id=int(row["driver_id"]),
                status="exported",
                source=str(row.get("source") or "auto"),
                odometer_start=row.get("odometer_start"),
                odometer_end=row.get("odometer_end"),
                fuel_start=row.get("fuel_start"),
                fuel_issued=row.get("fuel_issued"),
                fuel_end=row.get("fuel_end"),
                route_json=str(row.get("route_json") or "[]"),
                warnings_json=row.get("warnings_json"),
            )

        filename = (
            f"gsm_waybills_{period_from.isoformat()}_{period_to.isoformat()}.zip"
        )
        return buf.getvalue(), filename


def _as_date(value: Any) -> date:
    if isinstance(value, date):
        return value
    return date.fromisoformat(str(value)[:10])


def _parse_route_list(raw: Any) -> list[dict[str, Any]]:
    if raw is None or raw == "":
        return []
    try:
        parsed = json.loads(raw) if isinstance(raw, str) else raw
    except (TypeError, ValueError, json.JSONDecodeError):
        return []
    if not isinstance(parsed, list):
        return []
    return [item for item in parsed if isinstance(item, dict)]
