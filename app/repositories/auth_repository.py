from __future__ import annotations

import base64
import hashlib
import hmac
import os
import sqlite3
from typing import Any

from app.core.settings import get_settings


def _hash_password(password: str, *, salt: str | None = None) -> str:
    salt_bytes = os.urandom(16) if salt is None else base64.b64decode(salt.encode("ascii"))
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt_bytes, 200_000)
    return f"{base64.b64encode(salt_bytes).decode('ascii')}${base64.b64encode(digest).decode('ascii')}"


def _verify_password(password: str, stored_hash: str) -> bool:
    try:
        salt, expected = stored_hash.split("$", 1)
    except ValueError:
        return False
    candidate = _hash_password(password, salt=salt)
    return hmac.compare_digest(candidate, f"{salt}${expected}")


class AuthRepository:
    def __init__(self, db_path: str | None = None) -> None:
        settings = get_settings()
        self.db_path = db_path or str(settings.plita_db_path)
        self.settings = settings

    def init_schema(self) -> None:
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS app_users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    username TEXT NOT NULL UNIQUE,
                    password_hash TEXT NOT NULL,
                    role TEXT NOT NULL,
                    manager_id INTEGER,
                    is_active INTEGER NOT NULL DEFAULT 1,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                )
                """
            )
            cursor.execute(
                """
                CREATE INDEX IF NOT EXISTS idx_app_users_username
                ON app_users(username)
                """
            )
            conn.commit()

    def ensure_bootstrap_admin(self) -> None:
        self.init_schema()
        username = self.settings.bootstrap_admin_username
        password = self.settings.bootstrap_admin_password
        if not username or not password:
            return

        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT id FROM app_users WHERE username = ?", (username,))
            row = cursor.fetchone()
            password_hash = _hash_password(password)
            if row:
                cursor.execute(
                    """
                    UPDATE app_users
                    SET password_hash = ?, role = ?, is_active = 1
                    WHERE username = ?
                    """,
                    (password_hash, self.settings.bootstrap_admin_role, username),
                )
            else:
                cursor.execute(
                    """
                    INSERT INTO app_users(username, password_hash, role, is_active)
                    VALUES(?, ?, ?, 1)
                    """,
                    (username, password_hash, self.settings.bootstrap_admin_role),
                )
            conn.commit()

    def authenticate(self, username: str, password: str) -> dict[str, Any] | None:
        self.init_schema()
        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute(
                """
                SELECT id, username, password_hash, role, manager_id, is_active
                FROM app_users
                WHERE username = ?
                """,
                (username,),
            )
            row = cursor.fetchone()
            if not row:
                return None
            payload = dict(row)
            if not payload.get("is_active"):
                return None
            if not _verify_password(password, payload["password_hash"]):
                return None
            payload.pop("password_hash", None)
            return payload

    def list_users(self) -> list[dict[str, Any]]:
        self.init_schema()
        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute(
                "SELECT id, username, role, manager_id, is_active, created_at FROM app_users ORDER BY username"
            )
            return [dict(row) for row in cursor.fetchall()]

