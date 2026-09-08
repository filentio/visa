"""
Parser agent — превращает сырой пост в структуру вакансии.
Фиксы:
  - link: берём href из HTML-ссылок, не текст кнопки (избегаем плейсхолдеров)
  - recruiter_handle: валидируем — только @username без точек/слешей/двоеточий
  - Дополнительно: если ссылка ведёт на LinkedIn/HH/другой не-TG ресурс,
    кладём её в link, а recruiter_handle оставляем None
"""
from __future__ import annotations

import json
import logging
import re
import time
from typing import Optional

from jobsignal.db import get_session_factory, RawPost, Vacancy, VacancyStatus

log = logging.getLogger("jobsignal")

# ── TG handle validation ──────────────────────────────────────────────────────

# Valid TG handle: @word_chars only, 5+ chars, NO dots/slashes/colons
_TG_HANDLE_RE = re.compile(r"^@?([A-Za-z0-9_]{5,32})$")
# Patterns that are NOT TG handles (LinkedIn, HH, email, etc.)
_NOT_TG_PATTERNS = [
    r"linkedin\.com",
    r"linkedin:",
    r"hh\.ru",
    r"t\.me/joinchat",   # group invite links — not a personal handle
    r"@[^A-Za-z0-9_]",  # handle with non-word chars
    r"\.",               # dot in "handle" → domain or email
    r"/",                # slash → URL path
    r":",                # colon → URL scheme or LinkedIn shorthand
]
_NOT_TG_RE = re.compile("|".join(_NOT_TG_PATTERNS), re.IGNORECASE)


def _normalize_handle(raw: Optional[str]) -> Optional[str]:
    """
    Return clean @handle if it's a real TG username, else None.
    LinkedIn URLs, email-like strings, etc. → None.
    """
    if not raw:
        return None
    raw = raw.strip()
    # reject anything that looks like a URL or LinkedIn
    if _NOT_TG_RE.search(raw):
        return None
    # strip leading @
    clean = raw.lstrip("@")
    if _TG_HANDLE_RE.match(clean):
        return f"@{clean}"
    return None


# ── link extraction from raw HTML/text ───────────────────────────────────────

# Placeholder patterns the LLM sometimes returns instead of real URLs
_PLACEHOLDER_RE = re.compile(
    r"^\[.*?\]$|"           # [ссылка], [link], [url]
    r"^ссылка$|"
    r"^link$|"
    r"^url$|"
    r"^#|"                  # anchor-only
    r"^$",
    re.IGNORECASE,
)

def _is_real_link(link: Optional[str]) -> bool:
    if not link:
        return False
    link = link.strip()
    if _PLACEHOLDER_RE.match(link):
        return False
    if not link.startswith(("http://", "https://", "t.me/")):
        return False
    return True


def _extract_links_from_text(text: str) -> list[str]:
    """Pull all real URLs from raw post text."""
    urls = re.findall(
        r"https?://[^\s\)\]\"'<>]+|t\.me/[A-Za-z0-9_/]+",
        text,
    )
    return [u.rstrip(".,;)") for u in urls]


def _pick_best_link(links: list[str], post_url: Optional[str]) -> Optional[str]:
    """
    Priority:
    1. hh.ru vacancy link
    2. External apply link (not t.me channel link)
    3. Any t.me link that's not a channel post
    4. post_url as fallback
    """
    hh = [u for u in links if re.search(r"hh\.ru/vacancy/\d+", u)]
    if hh:
        return hh[0]

    external = [
        u for u in links
        if not re.search(r"t\.me/[A-Za-z0-9_]+/\d+", u)  # not a channel post
        and not u == post_url
    ]
    if external:
        return external[0]

    tme = [u for u in links if "t.me/" in u and u != post_url]
    if tme:
        return tme[0]

    return post_url


# ── LLM prompt ────────────────────────────────────────────────────────────────

SYSTEM = """Ты парсер вакансий. Извлекай поля из текста вакансии и возвращай ТОЛЬКО JSON без markdown.

Правила:
- role: название должности (строка или null)
- company: название компании (строка или null)  
- recruiter_handle: ТОЛЬКО личный Telegram @username рекрутёра (начинается с @, только буквы/цифры/подчёркивание, без точек/слешей). Если контакт — LinkedIn, email, HH или что-то другое — верни null.
- salary: зарплата строкой или null
- location: город/формат работы или null
- link: прямая ссылка для отклика (hh.ru, форма, сайт). НЕ возвращай текст вида [ссылка] или [link] — только реальный URL начиная с https:// или t.me/. Если ссылки нет — null.
- description: краткое описание обязанностей (1-2 предложения) или null
- is_vacancy: true если это объявление о вакансии, false если новость/анонс/болталка

Пример ответа:
{"role":"Product Manager","company":"Acme","recruiter_handle":"@hr_anna","salary":"от 200 000 ₽","location":"Москва/удалённо","link":"https://hh.ru/vacancy/123","description":"Развитие B2B-продукта","is_vacancy":true}"""


# ── Parser ────────────────────────────────────────────────────────────────────

class Parser:
    def __init__(self):
        self._sf = get_session_factory()

    def run(self, limit: Optional[int] = None) -> dict:
        from jobsignal.llm import call_llm, LLMError
        batch = int(__import__('os').environ.get("PARSER_BATCH_LIMIT", "200"))
        if limit:
            batch = limit

        session = self._sf()
        posts = (
            session.query(RawPost)
            .filter(RawPost.parsed == False, RawPost.text.isnot(None))
            .limit(batch)
            .all()
        )

        parsed = vacancies = skipped = errors = 0

        for post in posts:
            try:
                result = self._parse_post(post, session)
                if result == "vacancy":
                    vacancies += 1
                elif result == "skipped":
                    skipped += 1
                parsed += 1
                post.parsed = True
                session.commit()
            except Exception as exc:
                log.warning("[parser] post %d error: %s", post.id, exc)
                errors += 1
                session.rollback()

        session.close()
        summary = {
            "agent": "parser",
            "parsed": parsed,
            "vacancies": vacancies,
            "skipped": skipped,
            "errors": errors,
        }
        log.info("[parser] обработано: %d, вакансий: %d, отсеяно: %d, ошибок API: %d",
                 parsed, vacancies, skipped, errors)
        return summary

    def _parse_post(self, post: RawPost, session) -> str:
        from jobsignal.llm import call_llm, LLMError, _loads_lenient

        text = (post.text or "").strip()
        if not text or len(text) < 30:
            return "skipped"

        # Extract all URLs from raw text BEFORE sending to LLM
        text_links = _extract_links_from_text(text)

        prompt = f"Текст вакансии:\n\n{text[:3000]}"
        try:
            raw = call_llm(
                system=SYSTEM,
                user=prompt,
                max_tokens=400,
            )
        except LLMError as exc:
            log.warning("[parser] LLM error post %d: %s", post.id, exc)
            return "error"

        try:
            data = _loads_lenient(raw)
        except Exception:
            log.warning("[parser] bad JSON post %d: %.120s", post.id, raw)
            return "skipped"

        if not data.get("is_vacancy"):
            return "skipped"

        # ── link: prefer extracted links, use LLM link only if real ──────────
        llm_link = data.get("link")
        if not _is_real_link(llm_link):
            llm_link = None

        all_links = text_links[:]
        if llm_link and llm_link not in all_links:
            all_links.insert(0, llm_link)

        best_link = _pick_best_link(all_links, post.post_url)

        # ── recruiter_handle: validate strictly ───────────────────────────────
        raw_handle = data.get("recruiter_handle")
        clean_handle = _normalize_handle(raw_handle)

        # Also try to find @handle in text if LLM missed or returned garbage
        if not clean_handle:
            tg_in_text = re.findall(r"@([A-Za-z0-9_]{5,32})", text)
            for h in tg_in_text:
                candidate = f"@{h}"
                if _normalize_handle(candidate):
                    clean_handle = candidate
                    break

        # If the "handle" was actually a linkedin/external link, put it in link
        if raw_handle and not clean_handle and _is_real_link(raw_handle):
            if not best_link:
                best_link = raw_handle

        # ── detect contact type ───────────────────────────────────────────────
        from jobsignal.agents.outreach import detect_contact_type

        class _V:
            recruiter_handle = clean_handle
            link = best_link

        contact_type = detect_contact_type(_V())

        # ── save vacancy ──────────────────────────────────────────────────────
        v = Vacancy(
            raw_post_id=post.id,
            channel_id=post.channel_id,
            role=data.get("role"),
            company=data.get("company"),
            recruiter_handle=clean_handle,
            salary=data.get("salary"),
            location=data.get("location"),
            link=best_link,
            description=data.get("description"),
            status=VacancyStatus.new,
            contact_type=contact_type,
        )
        session.add(v)
        return "vacancy"
