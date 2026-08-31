"""Domain errors for shipment logistics (shared by facade, completion, repository)."""

from __future__ import annotations


class ShipmentError(ValueError):
    """Domain validation error for shipment operations (maps to 422)."""

    def __init__(self, message: str, *, code: str = "shipment_error") -> None:
        super().__init__(message)
        self.code = code
