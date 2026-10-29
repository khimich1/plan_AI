"""Nginx завода не режет загрузку прайса дефолтным 1m."""

from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

_PRICE_LOCATIONS = ("/api/v1/prices/", "/api/v1/nomenclature/")


def _location_body(text: str, prefix: str) -> str:
    match = re.search(rf"location {re.escape(prefix)} \{{(.*?)\n[ \t]*\}}", text, re.S)
    assert match is not None, prefix
    return match.group(1)


def test_plant_nginx_does_not_cap_price_uploads() -> None:
    for relative in (
        "frontend/nginx.app.locations.conf",
        "deploy/nginx-docker-ssl.conf",
    ):
        text = (ROOT / relative).read_text(encoding="utf-8")
        for prefix in _PRICE_LOCATIONS:
            body = _location_body(text, prefix)
            assert "client_max_body_size 0;" in body, relative
        assert "client_max_body_size 50m;" in _location_body(text, "/api/")

    host = (ROOT / "deploy/nginx-host-ssl.conf").read_text(encoding="utf-8")
    assert "client_max_body_size 0;" in host
