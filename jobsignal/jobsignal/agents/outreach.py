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
# Здесь был шаблон письма с зашитыми «достижениями» по профилям: «вывел продукт
# с 0 до 130k MAU», «работал в финтехе (Альфа, Финам)», «−22% отвала» —
# половины этого в резюме нет. Кнопка «Откликнуться» в дашборде показывала
# такой текст для hh, LinkedIn и форм. Убрано 23.10: письмо теперь одно на
# весь проект — notify_bot._cover_text по резюме профиля; текст передаёт
# вызывающий (см. dashboard.api_outreach_data).


# ── outreach data builder ─────────────────────────────────────────────────────

def get_outreach_data(vacancy, best_profile_key: str = "",
                      cover_text: str | None = None) -> dict:
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
        data["cover_text"] = cover_text
        data["action_label"] = "Открыть профиль в LinkedIn"
        data["link"] = linkedin_url or link
        data["instructions"] = (
            "Скопируй текст → открой профиль рекрутёра в LinkedIn → "
            "отправь сообщение или InMail с этим текстом."
        )

    elif contact_type == "hh":
        canonical = normalize_hh_link(link)
        data["link"] = canonical or link
        data["cover_text"] = cover_text
        data["action_label"] = "Открыть на hh.ru"
        data["instructions"] = (
            "Скопируй текст → открой вакансию → вставь в поле "
            "«Сопроводительное письмо» → приложи резюме → отправь."
        )

    elif contact_type == "form":
        data["cover_text"] = cover_text
        data["action_label"] = "Открыть форму"
        data["instructions"] = (
            "Скопируй текст → открой форму → вставь в поле «о себе» → "
            "приложи резюме → отправь."
        )

    elif contact_type == "unknown":
        data["action_label"] = "Открыть вакансию"
        data["instructions"] = "Ссылка на отклик не найдена — открой пост вакансии."

    return data
