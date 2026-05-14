# -*- coding: utf-8 -*-
"""Тесты резолвера марки бетона (pb_reinforcement_series)."""

from __future__ import annotations

import sqlite3
import tempfile
from pathlib import Path

import unittest

from core.concrete_grade_resolver import (
    enrich_orders_2d_concrete_grade,
    normalize_concrete_grade,
    resolve_concrete_grade,
)


class TestNormalizeConcreteGrade(unittest.TestCase):
    def test_latin_m(self) -> None:
        self.assertEqual(normalize_concrete_grade("M500"), "М500")

    def test_cyrillic_kept(self) -> None:
        self.assertEqual(normalize_concrete_grade("М400"), "М400")


class TestConcreteGradeSeries(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".db")
        self._tmp.close()
        self.db_path = Path(self._tmp.name)
        conn = sqlite3.connect(str(self.db_path))
        conn.execute(
            """
            CREATE TABLE pb_reinforcement_series (
                length_dm INTEGER,
                load_code INTEGER,
                reinforcement_value REAL,
                concrete_grade TEXT,
                PRIMARY KEY(length_dm, load_code)
            )
            """
        )
        conn.execute(
            "INSERT INTO pb_reinforcement_series VALUES (45, 8, 1.0, 'М400')"
        )
        conn.execute(
            "INSERT INTO pb_reinforcement_series VALUES (71, 8, 1.0, 'M500')"
        )
        conn.commit()
        conn.close()

    def tearDown(self) -> None:
        self.db_path.unlink(missing_ok=True)

    def test_explicit_wins_series(self) -> None:
        g = resolve_concrete_grade(
            concrete_grade_explicit="М500",
            plate_name="ПБ 45-12-8п",
            length_m=4.5,
            load_code=800,
            db_path=str(self.db_path),
        )
        self.assertEqual(g, "М500")

    def test_reads_series_by_length_dm(self) -> None:
        g = resolve_concrete_grade(
            concrete_grade_explicit=None,
            plate_name="",
            length_m=4.5,
            load_code=8,
            db_path=str(self.db_path),
        )
        self.assertEqual(g, "М400")

    def test_normalize_m500_from_series(self) -> None:
        g = resolve_concrete_grade(
            concrete_grade_explicit=None,
            plate_name="",
            length_m=7.1,
            load_code=8,
            db_path=str(self.db_path),
        )
        self.assertEqual(g, "М500")


class TestEnrichOrders2d(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".db")
        self._tmp.close()
        self.db_path = Path(self._tmp.name)
        conn = sqlite3.connect(str(self.db_path))
        conn.execute(
            """
            CREATE TABLE pb_reinforcement_series (
                length_dm INTEGER,
                load_code INTEGER,
                reinforcement_value REAL,
                concrete_grade TEXT,
                PRIMARY KEY(length_dm, load_code)
            )
            """
        )
        conn.execute(
            "INSERT INTO pb_reinforcement_series VALUES (60, 8, 1.0, 'М400')"
        )
        conn.commit()
        conn.close()

    def tearDown(self) -> None:
        self.db_path.unlink(missing_ok=True)

    def test_enrich_fills_blank(self) -> None:
        orders = [
            {"length": 6.0, "width": 1200, "qty": 1, "load_code": 800, "plate_name": "x"},
        ]
        enrich_orders_2d_concrete_grade(orders, db_path=str(self.db_path))
        self.assertEqual(orders[0].get("concrete_grade"), "М400")


if __name__ == "__main__":
    unittest.main()
