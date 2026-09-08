"""
HHCollector — сбор вакансий с hh.ru через публичный API (без авторизации).

Эндпоинт: GET https://api.hh.ru/vacancies
Не требует токена — анонимные запросы поддерживаются официально.
Результаты складываются в raw_posts с channel_id специального HH-канала,
дальше идёт тот же пайплайн: parser → dedup → matcher.
"""
from __future__ import annotations

import logging
import time
from datetime import datetime, timezone
from typing import Optional

import requests

from jobsignal.db import get_session_factory, Channel, RawPost

log = logging.getLogger("jobsignal")

HH_API = "https://api.hh.ru/vacancies"
HH_VACANCY_URL = "https://hh.ru/vacancy/{}"

HEADERS = {
    "User-Agent": "JobSignal/1.0 (personal job search tool)",
    "HH-User-Agent": "JobSignal/1.0 (personal job search tool)",
}

# Поисковые запросы под три профиля
DEFAULT_SEARCHES = [
    # CPO / Head of Product
    {"text": "CPO", "area": 1, "experience": "between3And6"},
    {"text": "Chief Product Officer", "area": 1},
    {"text": "Head of Product", "area": 1, "experience": "between3And6"},
    {"text": "Директор по продукту", "area": 1},
    # Senior AI PM
    {"text": "AI Product Manager", "area": 1},
    {"text": "ML Product Manager", "area": 1},
    {"text": "Product Manager AI", "area": 1},
    # Senior PM fintech
    {"text": "Product Manager fintech", "area": 1},
    {"text": "Продакт менеджер финтех", "area": 1},
    {"text": "Product Owner банк", "area": 1},
    # General senior PM
    {"text": "Senior Product Manager", "area": 1, "experience": "between3And6"},
    {"text": "Lead Product Manager", "area": 1},
]

# area=1 = Москва, area=2 = Санкт-Петербург, area=113 = Россия
# professional_roles: 96=Продуктовый менеджер, 104=Руководитель
PM_PROFESSIONAL_ROLES = [96, 104, 157]  # PM, IT-директор, Бизнес-аналитик


def _hh_channel(session) -> Channel:
    """Получить или создать виртуальный канал для hh.ru вакансий."""
    ch = session.query(Channel).filter_by(handle="hh_ru").first()
    if not ch:
        ch = Channel(
            handle="hh_ru",
            title="hh.ru (API)",
            niche="hh",
            active=True,
            source="hh",
        )
        session.add(ch)
        session.commit()
        session.refresh(ch)
    return ch


def _fetch_vacancies(params: dict, per_page: int = 50) -> list[dict]:
    """Fetch one page of vacancies from hh.ru API."""
    p = {**params, "per_page": per_page, "page": 0, "only_with_salary": False}
    try:
        r = requests.get(HH_API, params=p, headers=HEADERS, timeout=15)
        if r.status_code == 200:
            return r.json().get("items", [])
        log.warning("[hh] HTTP %s for %s", r.status_code, params.get("text"))
    except Exception as exc:
        log.warning("[hh] error for %s: %s", params.get("text"), exc)
    return []


def _format_post_text(item: dict) -> str:
    """Convert hh.ru vacancy item to text for parser."""
    parts = []

    name = item.get("name", "")
    if name:
        parts.append(name)

    employer = item.get("employer", {})
    if employer.get("name"):
        parts.append(f"Компания: {employer['name']}")

    salary = item.get("salary")
    if salary:
        frm = salary.get("from")
        to = salary.get("to")
        cur = salary.get("currency", "RUB")
        if frm and to:
            parts.append(f"Зарплата: {frm}–{to} {cur}")
        elif frm:
            parts.append(f"Зарплата: от {frm} {cur}")
        elif to:
            parts.append(f"Зарплата: до {to} {cur}")

    area = item.get("area", {})
    if area.get("name"):
        parts.append(f"Локация: {area['name']}")

    schedule = item.get("schedule", {})
    if schedule.get("name"):
        parts.append(f"График: {schedule['name']}")

    experience = item.get("experience", {})
    if experience.get("name"):
        parts.append(f"Опыт: {experience['name']}")

    snippet = item.get("snippet", {})
    if snippet.get("requirement"):
        req = snippet["requirement"].replace("<highlighttext>", "").replace("</highlighttext>", "")
        parts.append(f"Требования: {req}")
    if snippet.get("responsibility"):
        resp = snippet["responsibility"].replace("<highlighttext>", "").replace("</highlighttext>", "")
        parts.append(f"Обязанности: {resp}")

    url = HH_VACANCY_URL.format(item.get("id", ""))
    parts.append(f"Ссылка: {url}")

    return "\n".join(parts)


class HHCollector:
    def __init__(self, searches: Optional[list[dict]] = None):
        self._sf = get_session_factory()
        self.searches = searches or DEFAULT_SEARCHES

    def run(self) -> dict:
        session = self._sf()
        ch = _hh_channel(session)

        # existing hh vacancy IDs to avoid duplicates
        existing = {
            row[0]
            for row in session.query(RawPost.tg_message_id)
            .filter(RawPost.channel_id == ch.id)
            .all()
        }

        added = 0
        seen_ids: set[int] = set()

        for search_params in self.searches:
            items = _fetch_vacancies(search_params)
            log.info("[hh] '%s' → %d items", search_params.get("text"), len(items))

            for item in items:
                vac_id = int(item.get("id", 0))
                if not vac_id or vac_id in seen_ids or vac_id in existing:
                    continue
                seen_ids.add(vac_id)

                text = _format_post_text(item)
                published = item.get("published_at")
                try:
                    posted_at = datetime.fromisoformat(
                        published.replace("Z", "+00:00")
                    ) if published else datetime.now(timezone.utc)
                except Exception:
                    posted_at = datetime.now(timezone.utc)

                post_url = HH_VACANCY_URL.format(vac_id)

                post = RawPost(
                    channel_id=ch.id,
                    tg_message_id=vac_id,   # используем hh vacancy id как уникальный id
                    text=text,
                    post_url=post_url,
                    posted_at=posted_at,
                    parsed=False,
                )
                session.add(post)
                added += 1

            time.sleep(0.5)  # rate limit

        try:
            session.commit()
        except Exception as exc:
            session.rollback()
            log.error("[hh] db commit error: %s", exc)

        session.close()
        result = {"agent": "hh_collector", "added": added, "searches": len(self.searches)}
        log.info("[hh] добавлено вакансий: %d", added)
        return result
