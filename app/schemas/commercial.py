from __future__ import annotations

from pydantic import BaseModel, Field


class CommercialParseRequest(BaseModel):
    text: str = Field(min_length=1)


class CommercialPreviewRequest(BaseModel):
    text: str = Field(min_length=1)


class CommercialPreviewXlsxRequest(BaseModel):
    customer_name: str = Field(min_length=1)
    manager_name: str = Field(min_length=1)
    manager_phone: str = ""
    manager_email: str = ""
    discount_percent: float = Field(default=0, ge=0, le=100)
    delivery_conditions: str = ""
    payment_conditions: str = ""


class CommercialRecognizeScreenResponse(BaseModel):
    recognized_text: str
    normalized_text: str
    lines: list[str] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)
    method: str = ""
    confidence: float = 0.0


class CommercialPreviewCheckXlsxRequest(BaseModel):
    plates_text: str = Field(min_length=1)
    recognized_text: str = ""

