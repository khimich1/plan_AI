from pathlib import Path

from app.repositories.auth_repository import AuthRepository


def make_repository(tmp_path: Path) -> AuthRepository:
    db_path = tmp_path / "auth.db"
    return AuthRepository(str(db_path))


def test_create_user_and_authenticate(tmp_path: Path):
    repository = make_repository(tmp_path)

    user, created = repository.create_or_update_user(
        username="admin",
        password="StrongPassword123!",
        role="admin",
    )

    assert created is True
    assert user["username"] == "admin"
    assert user["role"] == "admin"

    authenticated_user = repository.authenticate("admin", "StrongPassword123!")

    assert authenticated_user is not None
    assert authenticated_user["username"] == "admin"
    assert authenticated_user["role"] == "admin"


def test_create_or_update_user_updates_password_and_role(tmp_path: Path):
    repository = make_repository(tmp_path)
    repository.create_or_update_user(
        username="operator",
        password="InitialPassword123!",
        role="manager",
    )

    user, created = repository.create_or_update_user(
        username="operator",
        password="UpdatedPassword123!",
        role="admin",
        is_active=False,
    )

    assert created is False
    assert user["username"] == "operator"
    assert user["role"] == "admin"
    assert user["is_active"] == 0
    assert repository.authenticate("operator", "InitialPassword123!") is None
    assert repository.authenticate("operator", "UpdatedPassword123!") is None


def test_get_user_by_username_can_include_password_hash(tmp_path: Path):
    repository = make_repository(tmp_path)
    repository.create_or_update_user(
        username="auditor",
        password="AuditPassword123!",
        role="admin",
    )

    public_user = repository.get_user_by_username("auditor")
    raw_user = repository.get_user_by_username("auditor", include_password_hash=True)

    assert public_user is not None
    assert raw_user is not None
    assert "password_hash" not in public_user
    assert "password_hash" in raw_user
