from __future__ import annotations

from typing import Protocol


class NomenclatureCacheFiller(Protocol):
    def __call__(self) -> None: ...
