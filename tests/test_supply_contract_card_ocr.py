"""Зрение карточки: один Extract и точечный повтор. Живой API не вызывается."""

from __future__ import annotations

import asyncio
import json
import sqlite3
from io import BytesIO
from pathlib import Path
from unittest.mock import MagicMock, patch

from docx import Document
from PIL import Image

from app.services.supply_contract_service import SupplyContractService
from core import kp_db_schema
from core.config.settings import get_settings
from core.kp_db_common import _connect
from core.ocr.prompts import CONTRACT_CARD_EXTRACT_PROMPT
from core.ocr.providers.gigachat import GigaChatProvider
from core.supply_contract_card_ocr import fill_card_holes, recognize_contract_card
from core.supply_contract_card_parse import parse_card_text

ADMIN = {"id": 7, "role": "admin", "username": "admin"}

_ACCEPTED_DIGITS = {
    "inn": "7604010011",
    "kpp": "760401001",
    "ogrn": "1027600000009",
    "bik": "044525225",
    "account": "40702810000000000007",
    "corr_account": "30101810000000000760",
}


def _fresh_db(tmp_path: Path) -> str:
    db_path = str(tmp_path / "supply.db")
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    return db_path


def _seed(db_path: str) -> None:
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        cur = conn.execute(
            """
            INSERT INTO counterparties (code_1c, name, name_normalized, inn, kpp)
            VALUES ('00-1', 'РОМАШКА ООО', 'ромашка ооо', '7701000001', '770101001')
            """
        )
        client_id = int(cur.lastrowid)
        conn.execute(
            """
            INSERT INTO KP_offers (
                kp_id, creation_date, customer_name, manager_name, counterparty_id
            ) VALUES (1, '2026-09-01', 'РОМАШКА ООО', 'Иван Иванов', ?)
            """,
            (client_id,),
        )
        conn.execute(
            "INSERT INTO kp_meta (kp_id, status, owner_user_id) VALUES (1, 'в архиве', 7)"
        )


def _png() -> bytes:
    buffer = BytesIO()
    Image.new("RGB", (8, 8), "white").save(buffer, format="PNG")
    return buffer.getvalue()


def _count(db_path: str) -> int:
    with _connect(db_path) as conn:
        return int(conn.execute("SELECT COUNT(*) FROM supply_contract").fetchone()[0])


def test_extract_prompt_requires_header_and_verbatim_digits() -> None:
    prompt = CONTRACT_CARD_EXTRACT_PROMPT
    assert "ИНН КПП ОГРН" in prompt
    assert "не сжима" in prompt
    assert "БИК" in prompt
    assert '"okved"' in prompt


def test_retry_does_not_overwrite_an_accepted_inn() -> None:
    calls: list[str] = []

    class Provider:
        async def extract_contract_fields(self, **kwargs: object) -> tuple[dict, float]:
            calls.append("extract")
            return {"inn": "7604010011", "email": "a@b.ru", "okved": "23.61"}, 0.0

        async def reread_contract_fields(self, **kwargs: object) -> tuple[dict, float]:
            calls.append(str(kwargs.get("user_text", "")))
            return {"inn": "7707083893", "bik": "044525225", "okved": "52.29"}, 0.0

        async def verify_contract_fields(self, **kwargs: object) -> tuple[dict, float]:
            raise AssertionError("полный Verify по карточке не вызывается")

    result = asyncio.run(recognize_contract_card(_png(), filename="card.png", provider=Provider()))

    assert result.fields["inn"] == "7604010011"
    assert result.fields["bik"] == "044525225"
    assert calls[0] == "extract"
    assert len(calls) == 2
    assert "inn" not in calls[1]
    assert "bik" in calls[1]
    assert "okved" not in calls[1]
    assert "corrections" not in calls[1]
    assert result.fields["okved"] == "23.61"


def test_retry_is_skipped_when_every_digit_field_is_accepted() -> None:
    class Provider:
        async def extract_contract_fields(self, **kwargs: object) -> tuple[dict, float]:
            return dict(_ACCEPTED_DIGITS), 0.0

        async def reread_contract_fields(self, **kwargs: object) -> tuple[dict, float]:
            raise AssertionError("повтор не нужен, все цифровые поля приняты")

        async def verify_contract_fields(self, **kwargs: object) -> tuple[dict, float]:
            raise AssertionError("полный Verify по карточке не вызывается")

    result = asyncio.run(recognize_contract_card(_png(), filename="card.png", provider=Provider()))

    assert result.fields["inn"] == "7604010011"
    assert result.fields["account"] == "40702810000000000007"
    assert result.verify_failed is False


def test_model_charter_wording_becomes_list_value() -> None:
    class Provider:
        async def extract_contract_fields(self, **kwargs: object) -> tuple[dict, float]:
            return {
                **_ACCEPTED_DIGITS,
                "authority_basis": "Устава",
                "signatory_verb": "действует",
                "legal_form": "ООО",
            }, 0.0

        async def reread_contract_fields(self, **kwargs: object) -> tuple[dict, float]:
            raise AssertionError("повтор не нужен, все цифровые поля приняты")

    result = asyncio.run(recognize_contract_card(_png(), filename="card.png", provider=Provider()))

    assert result.fields["authority_basis"] == "устав"
    assert result.fields["signatory_verb"] is None
    assert result.fields["legal_form"] is None


def test_png_with_ocr_provider_gigachat_uses_mock_not_openai(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.setenv("OCR_PROVIDER", "gigachat")
    monkeypatch.setenv("GIGACHAT_CREDENTIALS", "dGVzdA==")
    monkeypatch.setenv("GIGACHAT_MODEL", "GigaChat-2-Max")
    get_settings.cache_clear()

    card_json = json.dumps(
        {
            "legal_form": "ooo",
            "full_name": "Общество с ограниченной ответственностью «Ромашка»",
            "short_name": "Ромашка",
            "signatory_name": "Иванов Иван Иванович",
            "email": "is-ag @mail.ru",
            **_ACCEPTED_DIGITS,
        },
        ensure_ascii=False,
    )
    mock_client = MagicMock()
    mock_client.upload_file.return_value = MagicMock(id_="file-card")
    mock_client.chat.return_value = MagicMock(
        choices=[MagicMock(message=MagicMock(content=card_json))],
        usage=MagicMock(total_tokens=800),
    )

    def _boom_openai(*_a, **_k):
        raise AssertionError("OpenAI не должен вызываться при OCR_PROVIDER=gigachat")

    db_path = _fresh_db(tmp_path)
    _seed(db_path)
    service = SupplyContractService(db_path=db_path)

    with (
        patch(
            "core.ocr.pipeline.create_ocr_provider",
            return_value=(
                GigaChatProvider(settings=get_settings(), client=mock_client),
                "gigachat",
                "GigaChat-2-Max",
            ),
        ),
        patch("core.ocr.providers.openai.require_openai_client", side_effect=_boom_openai),
        patch(
            "core.ocr.parsing.parse_gpt_response",
            side_effect=AssertionError("парсер марок не должен вызываться"),
        ),
    ):
        result = asyncio.run(service.parse_for_kp(1, _png(), filename="card.png", user=ADMIN))

    assert mock_client.chat.call_count == 1
    sent = mock_client.chat.call_args.args[0].messages[0].content
    assert "ИНН КПП ОГРН" in sent
    assert "corrections" not in sent
    assert result["fields"]["inn"] == "7604010011"
    assert result["fields"]["email"] == "is-ag@mail.ru"
    assert _count(db_path) == 0
    get_settings.cache_clear()


_CLEAN_CARD = """ООО «Ромашка»
Общество с ограниченной ответственностью «Ромашка»
ИНН 7604010011
КПП 760401001
ОГРН 1027600000009
Генеральный директор Иванов Иван Иванович
Юридический адрес: г. Ярославль, ул. Ленина, 1
E-mail: is-ag@mail.ru
Банк: ПАО «Тест Банк»
Р/с 40702810000000000007
К/с 30101810000000000760
БИК 044525225
"""

_HOLE_CARD = """ООО «Ромашка»
Общество с ограниченной ответственностью «Ромашка»
ИНН 7604010011
КПП 760401001
ОГРН 1027600000009
Юридический адрес: (в соответствии с учредительными документами)
E-mail: is-ag@mail.ru
Банк: ПАО «Тест Банк»
Р/с 40702810000000000007
К/с 30101810000000000760
БИК 044525225
"""


def test_clean_text_does_not_call_the_provider() -> None:
    class Provider:
        async def complete_text_json(self, *, user_text: str) -> tuple[dict, float]:
            raise AssertionError("чистая карточка не должна уходить в чат")

    parsed = parse_card_text(_CLEAN_CARD)
    result = asyncio.run(fill_card_holes(parsed, _CLEAN_CARD, Provider()))

    assert result.fields["inn"] == "7604010011"
    assert result.fields["signatory_name"] == "Иванов Иван Иванович"


def test_hole_calls_one_text_chat_and_keeps_a_checked_inn() -> None:
    calls: list[str] = []

    class Provider:
        async def complete_text_json(self, *, user_text: str) -> tuple[dict, float]:
            calls.append(user_text)
            return {
                "inn": "7707083893",
                "signatory_name": "Иванов Иван Иванович",
                "signatory_position": "Директор",
                "legal_address": "г. Ярославль, ул. Ленина, 1",
            }, 0.0

    parsed = parse_card_text(_HOLE_CARD)
    result = asyncio.run(fill_card_holes(parsed, _HOLE_CARD, Provider()))

    assert len(calls) == 1
    assert "7604010011" in calls[0]
    assert "signatory_name" in calls[0]
    assert "0" in calls[0] and "7" in calls[0]
    assert result.fields["inn"] == "7604010011"
    assert result.fields["signatory_name"] == "Иванов Иван Иванович"
    assert result.fields["legal_address"] == "г. Ярославль, ул. Ленина, 1"
    assert "signatory_name" in result.doubtful
    assert "signatory_position" in result.doubtful
    assert "legal_address" in result.doubtful


def test_model_inn_that_fails_control_stays_out_of_fields() -> None:
    class Provider:
        async def complete_text_json(self, *, user_text: str) -> tuple[dict, float]:
            return {"inn": "7604010010"}, 0.0

    parsed = parse_card_text("ИНН 7604010010")
    result = asyncio.run(fill_card_holes(parsed, "ИНН 7604010010", Provider()))

    assert result.fields["inn"] is None
    assert "inn" in result.doubtful


def test_provider_error_keeps_local_fields_without_the_suspicious_address() -> None:
    class Provider:
        async def complete_text_json(self, *, user_text: str) -> tuple[dict, float]:
            raise RuntimeError("сеть недоступна")

    parsed = parse_card_text(_HOLE_CARD)
    result = asyncio.run(fill_card_holes(parsed, _HOLE_CARD, Provider()))

    assert result.fields["inn"] == "7604010011"
    assert result.fields["account"] == "40702810000000000007"
    assert result.fields["legal_address"] is None
    assert "legal_address" in result.doubtful


def test_two_inns_can_be_replaced_and_stay_doubtful() -> None:
    class Provider:
        async def complete_text_json(self, *, user_text: str) -> tuple[dict, float]:
            return {"inn": "7707083893"}, 0.0

    text = "ИНН 7604010011\nИНН 7707083893"
    parsed = parse_card_text(text)
    result = asyncio.run(fill_card_holes(parsed, text, Provider()))

    assert result.fields["inn"] == "7707083893"
    assert "inn" in result.doubtful


def test_unpaired_banks_become_bundles_from_one_text_reply() -> None:
    calls: list[str] = []

    class Provider:
        async def complete_text_json(self, *, user_text: str) -> tuple[dict, float]:
            calls.append(user_text)
            return {
                "banks": [
                    {
                        "bank_name": "Альфа",
                        "account": "40702810000000000007",
                        "corr_account": "30101810000000000760",
                        "bik": "044525225",
                    },
                    {
                        "bank_name": None,
                        "account": "40702810000000000010",
                        "corr_account": "30101810400000000225",
                        "bik": "044525225",
                    },
                ]
            }, 0.0

    text = (
        "К/с 30101810000000000760\n"
        "К/с 30101810400000000225\n"
        "Р/с 40702810000000000007\n"
        "Р/с 40702810000000000010\n"
        "БИК 044525225"
    )
    parsed = parse_card_text(text)
    result = asyncio.run(fill_card_holes(parsed, text, Provider()))

    assert len(calls) == 1
    assert "banks" in calls[0]
    assert len(result.banks) == 2
    assert result.banks[1]["bank_name"] is None
    assert result.fields["account"] is None
    assert result.fields["corr_account"] is None
    assert result.fields["bik"] is None
    assert result.fields["bank_name"] is None


def _docx(paragraphs: list[str]) -> bytes:
    document = Document()
    for paragraph in paragraphs:
        document.add_paragraph(paragraph)
    buffer = BytesIO()
    document.save(buffer)
    return buffer.getvalue()


def test_service_calls_text_provider_only_when_the_card_has_a_hole(tmp_path: Path) -> None:
    calls: list[str] = []

    class Provider:
        async def complete_text_json(self, *, user_text: str) -> tuple[dict, float]:
            calls.append(user_text)
            return {
                "signatory_name": "Иванов Иван Иванович",
                "signatory_position": "Директор",
                "legal_address": "г. Ярославль, ул. Ленина, 1",
            }, 0.0

    db_path = _fresh_db(tmp_path)
    _seed(db_path)
    service = SupplyContractService(db_path=db_path, card_text_provider=Provider())
    clean = _docx(_CLEAN_CARD.splitlines())
    hole = _docx(_HOLE_CARD.splitlines())

    clean_result = asyncio.run(service.parse_for_kp(1, clean, filename="clean.docx", user=ADMIN))
    assert clean_result["fields"]["inn"] == "7604010011"
    assert calls == []

    hole_result = asyncio.run(service.parse_for_kp(1, hole, filename="hole.docx", user=ADMIN))

    assert len(calls) == 1
    assert "image_base64" not in calls[0]
    assert hole_result["fields"]["inn"] == "7604010011"
    assert hole_result["fields"]["legal_address"] == "г. Ярославль, ул. Ленина, 1"
    assert hole_result["source_text"].startswith("ООО")
    assert _count(db_path) == 0
