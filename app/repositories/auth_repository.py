from __future__ import annotations

import base64
import hashlib
import hmac
import os
import sqlite3
from typing import Any

from app.core.settings import get_settings
from app.security.password_policy import validate_password


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

    def _row_to_payload(self, row: sqlite3.Row | None, *, include_password_hash: bool = False) -> dict[str, Any] | None:
        if row is None:
            return None
        payload = dict(row)
        if not include_password_hash:
            payload.pop("password_hash", None)
        return payload

    def get_user_by_id(self, user_id: int) -> dict[str, Any] | None:
        self.init_schema()
        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute(
                """
                SELECT id, username, role, manager_id, is_active, created_at
                FROM app_users
                WHERE id = ?
                """,
                (user_id,),
            )
            row = cursor.fetchone()
            return self._row_to_payload(row)

    def get_user_by_username(self, username: str, *, include_password_hash: bool = False) -> dict[str, Any] | None:
        self.init_schema()
        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute(
                """
                SELECT id, username, password_hash, role, manager_id, is_active, created_at
                FROM app_users
                WHERE username = ?
                """,
                (username,),
            )
            row = cursor.fetchone()
            return self._row_to_payload(row, include_password_hash=include_password_hash)

    def create_or_update_user(
        self,
        *,
        username: str,
        password: str,
        role: str,
        manager_id: int | None = None,
        is_active: bool = True,
    ) -> tuple[dict[str, Any], bool]:
        self.init_schema()
        password_hash = _hash_password(password)
        normalized_username = username.strip()
        if not normalized_username:
            raise ValueError("Username must not be empty.")
        if not password:
            raise ValueError("Password must not be empty.")
        validate_password(password)
        if not role.strip():
            raise ValueError("Role must not be empty.")

        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute("SELECT id FROM app_users WHERE username = ?", (normalized_username,))
            row = cursor.fetchone()
            if row is None:
                cursor.execute(
                    """
                    INSERT INTO app_users(username, password_hash, role, manager_id, is_active)
                    VALUES(?, ?, ?, ?, ?)
                    """,
                    (normalized_username, password_hash, role.strip(), manager_id, int(is_active)),
                )
                created = True
            else:
                cursor.execute(
                    """
                    UPDATE app_users
                    SET password_hash = ?, role = ?, manager_id = ?, is_active = ?
                    WHERE username = ?
                    """,
                    (password_hash, role.strip(), manager_id, int(is_active), normalized_username),
                )
                created = False
            conn.commit()
        user = self.get_user_by_username(normalized_username)
        if user is None:
            raise RuntimeError("User was not persisted.")
        return user, created

    def authenticate(self, username: str, password: str) -> dict[str, Any] | None:
        payload = self.get_user_by_username(username, include_password_hash=True)
        if not payload:
            return None
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

