from __future__ import annotations

from pydantic import BaseModel, Field

COUNTERPARTY_NAME_MAX_LENGTH = 255
COUNTERPARTY_CODE_1C_MAX_LENGTH = 64
COUNTERPARTY_INN_MAX_LENGTH = 12
COUNTERPARTY_KPP_MAX_LENGTH = 9
COUNTERPARTY_SEARCH_Q_MAX_LENGTH = 128


class CounterpartyShort(BaseModel):
    id: int
    code_1c: str
    name: str
    inn: str | None = None
    kpp: str | None = None


class CounterpartySearchResponse(BaseModel):
    items: list[CounterpartyShort]
    count: int


class CounterpartyCreateRequest(BaseModel):
    name: str = Field(min_length=1, max_length=COUNTERPARTY_NAME_MAX_LENGTH)
    code_1c: str = Field(min_length=1, max_length=COUNTERPARTY_CODE_1C_MAX_LENGTH)
    inn: str | None = Field(default=None, max_length=COUNTERPARTY_INN_MAX_LENGTH)
    kpp: str | None = Field(default=None, max_length=COUNTERPARTY_KPP_MAX_LENGTH)
    is_client: bool = True


class CounterpartyCreateResponse(BaseModel):
    item: CounterpartyShort
    warning: str | None = None
