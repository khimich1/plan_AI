#!/usr/bin/env python3
"""
Create or update a web admin user in plita.db.

Usage:
    python scripts/create_admin.py --username admin
    python scripts/create_admin.py --username admin --update-existing
"""

from __future__ import annotations

import argparse
import getpass
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from app.repositories.auth_repository import AuthRepository


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Create or update a web admin user.")
    parser.add_argument("--username", required=True, help="Username for the web admin.")
    parser.add_argument("--role", default="admin", help="Role to assign. Defaults to admin.")
    parser.add_argument(
        "--db-path",
        default=None,
        help="Optional path to the auth database. Defaults to settings.plita_db_path.",
    )
    parser.add_argument(
        "--update-existing",
        action="store_true",
        help="Allow overwriting password and role for an existing user.",
    )
    return parser.parse_args()


def prompt_password() -> str:
    password = getpass.getpass("Password: ")
    if not password:
        raise ValueError("Password must not be empty.")
    confirmation = getpass.getpass("Confirm password: ")
    if password != confirmation:
        raise ValueError("Passwords do not match.")
    return password


def main() -> int:
    args = parse_args()
    repository = AuthRepository(args.db_path)
    existing_user = repository.get_user_by_username(args.username)
    if existing_user and not args.update_existing:
        print(
            f"User '{args.username}' already exists. "
            "Re-run with --update-existing to reset the password and role."
        )
        return 1

    try:
        password = prompt_password()
        user, created = repository.create_or_update_user(
            username=args.username,
            password=password,
            role=args.role,
        )
    except ValueError as exc:
        print(f"Error: {exc}")
        return 1

    action = "created" if created else "updated"
    print(f"User '{user['username']}' {action} with role '{user['role']}'.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
