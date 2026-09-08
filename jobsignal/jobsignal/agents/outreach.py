"""
Outreach helper — Модуль A:
  Определяет тип контакта вакансии (tg / hh / form / unknown)
  и готовит данные для semi-auto отклика:
    - tg:   черновик сообщения (существующий composer)
    - hh:   нормализует hh.ru ссылку + сопроводительный текст
    - form: ссылка на форму + текст для вставки
"""
from __future__ import annotations

import re
import logging
from typing import Optional

log = logging.getLogger("jobsignal")

# ── contact type detection ────────────────────────────────────────────────────

HH_PATTERNS = [
    r"hh\.ru/vacancy/\d+",
    r"headhunter\.ru/vacancy/\d+",
    r"hh\.ru/applicant/vacancy\?vacancyId=\d+",
]

TG_HANDLE_RE = re.compile(r"^@[A-Za-z0-9_]{5,32}$")  # strict: only valid TG handles
TG_URL_RE = re.compile(r"t\.me/([A-Za-z0-9_]{5,})")

LINKEDIN_RE = re.compile(r"linkedin\.com|linkedin:", re.IGNORECASE)
EMAIL_RE = re.compile(r"^[^@]+@[^@]+\.[^@]+$")

FORM_PATTERNS = [
    r"docs\.google\.com/forms",
    r"forms\.gle/",
    r"typeform\.com",
    r"tally\.so",
    r"airtable\.com",
    r"notion\.site",
    r"greenhouse\.io/jobs",
    r"lever\.co/",
    r"huntflow\.ru",
    r"bamboohr\.com",
    r"workable\.com",
]


def detect_contact_type(vacancy) -> str:
    """
    Returns: 'tg' | 'hh' | 'form' | 'linkedin' | 'unknown'
    Priority: tg > hh > linkedin > form > unknown
    """
    handle = (vacancy.recruiter_handle or "").strip()

    # Only count handle as TG if it's a valid username (no dots/slashes/colons)
    if handle and TG_HANDLE_RE.match(handle):
        return "tg"

    link = (vacancy.link or "").lower()

    # LinkedIn via handle field or link
    if LINKEDIN_RE.search(handle) or LINKEDIN_RE.search(link):
        return "linkedin"

    # check hh
    for pat in HH_PATTERNS:
        if re.search(pat, link):
            return "hh"

    # check forms / ats
    for pat in FORM_PATTERNS:
        if re.search(pat, link):
            return "form"

    # tg link in vacancy link
    if re.search(r"t\.me/", link):
        return "tg"

    if link:
        return "form"  # has some link but not classified — treat as generic form

    return "unknown"


def normalize_hh_link(link: str) -> Optional[str]:
    """Extract canonical hh.ru vacancy URL."""
    m = re.search(r"(https?://(?:www\.)?hh\.ru/vacancy/\d+)", link)
    if m:
        return m.group(1)
    m = re.search(r"hh\.ru/vacancy/(\d+)", link)
    if m:
        return f"https://hh.ru/vacancy/{m.group(1)}"
    return link or None


# ── cover letter text for non-TG ─────────────────────────────────────────────

COVER_TEMPLATE = """Меня заинтересовала позиция {role}{company_part}.

{profile_highlight}

{interest_line}

Резюме прикладываю. Буду рад обсудить — когда удобно созвониться?"""

PROFILE_HIGHLIGHTS = {
    "cpo": (
        "Опыт — CPO/Head of Product: вывел продукт с 0 до 130k MAU, "
        "управлял портфелем из 5 продуктов и командой 15 человек, "
        "P&L-ответственность."
    ),
    "ai_pm": (
        "Специализация — AI Product: строил RAG-ассистент поддержки "
        "(−40% нагрузки на КЦ), LLM+OCR скоринг, рекомендательную систему. "
        "Self-built MVP на Cursor/Claude."
    ),
    "pm": (
        "Опыт — Senior PM: 20+ A/B-тестов в год, онбординг 15→3 мин, "
        "−22% отвала на активации, работал в финтехе (Альфа, Финам)."
    ),
}


def build_cover_letter(
    vacancy,
    profile_key: str,
    style: str = "default",
) -> str:
    """
    Build cover letter text for hh.ru / form apply.
    profile_key: cpo / ai_pm / pm
    """
    role = vacancy.role or "Product Manager"
    company_part = f" в {vacancy.company}" if vacancy.company else ""
    highlight = PROFILE_HIGHLIGHTS.get(
        profile_key, PROFILE_HIGHLIGHTS["pm"]
    )

    if style == "short":
        # 2-3 lines for forms with character limits
        return (
            f"Senior Product Manager, финтех + AI. "
            f"{highlight.split('.')[0]}. "
            f"Интересует {role}{company_part}. "
            "Резюме прикладываю."
        )

    interest_line = f"Позиция {role} совпадает с моим фокусом и опытом."

    return COVER_TEMPLATE.format(
        role=role,
        company_part=company_part,
        profile_highlight=highlight,
        interest_line=interest_line,
    ).strip()


# ── outreach data builder ─────────────────────────────────────────────────────

def get_outreach_data(vacancy, best_profile_key: str) -> dict:
    """
    Returns everything dashboard needs to render the apply block
    for a non-TG vacancy.
    """
    contact_type = vacancy.contact_type or detect_contact_type(vacancy)
    link = vacancy.link or ""

    # extract linkedin URL from handle if it was stored there
    linkedin_url = None
    handle_raw = getattr(vacancy, "recruiter_handle", "") or ""
    if LINKEDIN_RE.search(handle_raw):
        linkedin_url = handle_raw if handle_raw.startswith("http") else None

    data: dict = {
        "contact_type": contact_type,
        "link": link,
        "linkedin_url": linkedin_url,
        "cover_text": None,
        "action_label": "Открыть вакансию",
        "instructions": "",
    }

    if contact_type == "linkedin":
        data["cover_text"] = build_cover_letter(vacancy, best_profile_key, style="short")
        data["action_label"] = "Открыть профиль в LinkedIn"
        data["link"] = linkedin_url or link
        data["instructions"] = (
            "Скопируй текст → открой профиль рекрутёра в LinkedIn → "
            "отправь сообщение или InMail с этим текстом."
        )

    elif contact_type == "hh":
        canonical = normalize_hh_link(link)
        data["link"] = canonical or link
        data["cover_text"] = build_cover_letter(vacancy, best_profile_key)
        data["action_label"] = "Открыть на hh.ru"
        data["instructions"] = (
            "Скопируй текст → открой вакансию → вставь в поле "
            "«Сопроводительное письмо» → приложи резюме → отправь."
        )

    elif contact_type == "form":
        data["cover_text"] = build_cover_letter(
            vacancy, best_profile_key, style="short"
        )
        data["action_label"] = "Открыть форму"
        data["instructions"] = (
            "Скопируй текст → открой форму → вставь в поле «о себе» → "
            "приложи резюме → отправь."
        )

    elif contact_type == "unknown":
        data["action_label"] = "Открыть вакансию"
        data["instructions"] = "Ссылка на отклик не найдена — открой пост вакансии."

    return data
