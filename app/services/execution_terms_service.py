from __future__ import annotations

import re
from datetime import datetime, timedelta


class ExecutionTermsService:
    def normalize(self, raw_input: str) -> str:
        value = raw_input.strip()
        if not value:
            raise ValueError("Укажите срок изготовления.")

        deadline = self._parse_date(value)
        if deadline is None:
            raise ValueError("Укажите срок в формате ДД.ММ.ГГГГ, ГГГГ-ММ-ДД, N дней или N недель.")
        return deadline.strftime("%d.%m.%Y")

    def _parse_date(self, value: str) -> datetime | None:
        for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(value, fmt)
            except ValueError:
                continue

        match_days = re.search(r"(\d+)\s*(?:дн|день|дней|day|days)", value, re.IGNORECASE)
        if match_days:
            return datetime.now() + timedelta(days=int(match_days.group(1)))

        match_weeks = re.search(r"(\d+)\s*(?:нед|недел|недели|week|weeks)", value, re.IGNORECASE)
        if match_weeks:
            return datetime.now() + timedelta(weeks=int(match_weeks.group(1)))

        return None
