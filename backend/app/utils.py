from __future__ import annotations

import secrets
import string
from datetime import date, timedelta
from typing import Optional


def new_id() -> str:
    # UUID without dependency; good enough for MVP uniqueness.
    # 32 hex chars + 4 dashes (36 chars).
    import uuid

    return str(uuid.uuid4())


def random_start_date_within_last_6_months(today: Optional[date] = None) -> date:
    if today is None:
        today = date.today()
    delta_days = secrets.randbelow(180 + 1)
    return today - timedelta(days=delta_days)


def generate_contract_number(*, year: Optional[int] = None) -> str:
    if year is None:
        year = date.today().year
    # 6 digits
    n = secrets.randbelow(900000) + 100000
    return f"{n}/{year}"


_MRZ_MAP = {
    "RUS": "RUSSIA",
    "USA": "USA",
    "ARE": "UAE",
    "GBR": "UK",
}


def issuing_country_from_mrz(mrz: str) -> Optional[str]:
    """
    Best-effort MRZ parser to extract issuing country (3-letter code) from a passport MRZ.
    Typical: line1 starts with 'P<' + 3-letter issuing country.
    """
    raw = (mrz or "").strip()
    if not raw:
        return None
    lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]
    if not lines:
        return None
    line1 = lines[0]
    if len(line1) >= 5 and line1.startswith("P<"):
        code = line1[2:5]
        if code.isalpha():
            return code.upper()
    # fallback: find 'P<' anywhere
    idx = line1.find("P<")
    if idx >= 0 and len(line1) >= idx + 5:
        code = line1[idx + 2 : idx + 5]
        if code.isalpha():
            return code.upper()
    return None


def country_display_from_issuing(code: Optional[str]) -> str:
    if not code:
        return "RUSSIA, Moscow"
    code_u = code.upper()
    name = _MRZ_MAP.get(code_u, code_u)
    # Keep the example format used in runner payloads.
    if name == "RUSSIA":
        return "RUSSIA, Moscow"
    return name

