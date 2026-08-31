"""Smoke tests for ArchiveService.build_current_plan_gantt (audit Q2 v2)."""

from __future__ import annotations

import asyncio
from pathlib import Path
from unittest.mock import MagicMock

import pytest

from app.services.archive_service import ArchiveService


def test_build_current_plan_gantt_resolves_create_gantt_excel(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    out_file = tmp_path / "gantt.xlsx"
    out_file.write_bytes(b"xlsx")

    gantt_data = {
        "all_tracks": [],
        "plate_lookup_exact": {},
        "plate_lookup_by_length": {},
        "earliest_start_date": "2026-01-01",
    }

    monkeypatch.setattr(
        "app.services.archive_service.PlanDistributionService.get_all_plans_gantt_data",
        lambda self, repo: gantt_data,
    )

    called: dict = {}

    def fake_create_gantt_excel(**kwargs):
        called.update(kwargs)
        return str(out_file)

    monkeypatch.setattr(
        "app.services.archive_service.create_gantt_excel", fake_create_gantt_excel
    )

    repo = MagicMock()
    service = ArchiveService(repository=repo, outputs_dir=tmp_path)
    path = asyncio.run(service.build_current_plan_gantt())
    assert path == out_file
    assert called.get("output_dir") == str(tmp_path)
    assert called.get("start_date") == "2026-01-01"


def test_create_gantt_excel_imported_from_core() -> None:
    from app.services import archive_service
    from core.gantt_excel import create_gantt_excel as core_fn

    assert archive_service.create_gantt_excel is core_fn
