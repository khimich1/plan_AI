from app.security.session import create_session_token, decode_session_token


def test_session_roundtrip():
    token = create_session_token({"id": 1, "username": "demo", "role": "admin"}, ttl_seconds=60)

    payload = decode_session_token(token)

    assert payload is not None
    assert payload["id"] == 1
    assert payload["username"] == "demo"
    assert payload["role"] == "admin"

