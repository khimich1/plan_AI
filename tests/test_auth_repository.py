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


def test_get_user_by_id_returns_public_fields(tmp_path: Path):
    repository = make_repository(tmp_path)
    user, _created = repository.create_or_update_user(
        username="lookup",
        password="LookupPassword123!",
        role="manager",
    )

    loaded = repository.get_user_by_id(user["id"])

    assert loaded is not None
    assert loaded["username"] == "lookup"
    assert loaded["role"] == "manager"
    assert "password_hash" not in loaded
    assert repository.get_user_by_id(9999) is None


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


def test_list_users_paginates_by_username(tmp_path: Path):
    repository = make_repository(tmp_path)
    for index in range(5):
        repository.create_or_update_user(
            username=f"user_{index:02d}",
            password="PaginatePassword12!",
            role="manager",
        )

    first_page = repository.list_users(limit=2, offset=0)
    second_page = repository.list_users(limit=2, offset=2)
    last_page = repository.list_users(limit=2, offset=4)

    assert [user["username"] for user in first_page] == ["user_00", "user_01"]
    assert [user["username"] for user in second_page] == ["user_02", "user_03"]
    assert [user["username"] for user in last_page] == ["user_04"]
    assert repository.count_users() == 5


def test_bump_session_version_invalidates_existing_tokens(tmp_path: Path) -> None:
    repository = make_repository(tmp_path)
    user, _created = repository.create_or_update_user(
        username="revoke",
        password="RevokePassword123!",
        role="admin",
    )

    assert user["session_version"] == 0
    new_version = repository.bump_session_version(user["id"])
    loaded = repository.get_user_by_id(user["id"])

    assert new_version == 1
    assert loaded is not None
    assert loaded["session_version"] == 1


def test_get_users_page_returns_total_and_window(tmp_path: Path):
    repository = make_repository(tmp_path)
    repository.create_or_update_user(
        username="alpha",
        password="PaginatePassword12!",
        role="admin",
    )
    repository.create_or_update_user(
        username="beta",
        password="PaginatePassword12!",
        role="manager",
    )

    page = repository.get_users_page(limit=1, offset=1)

    assert page["total"] == 2
    assert page["limit"] == 1
    assert page["offset"] == 1
    assert len(page["items"]) == 1
    assert page["items"][0]["username"] == "beta"
