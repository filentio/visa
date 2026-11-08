"""Свежесть вакансии — по дате ПУБЛИКАЦИИ, а не по дате попадания в базу.

Зачем. До 08.11 возраст считался по vacancies.created_at — когда вакансия
легла в нашу базу. Пост телеграм-канала полугодовой давности, собранный
вчера, выглядел вчерашним. Дата публикации у источников есть:
raw_posts.posted_at — время поста в канале, published_at вакансии hh и
площадок (Хабр, getmatch). Нет её — берём created_at.

VACANCY_MAX_AGE_DAYS (по умолчанию 30, не больше 45): старше — не оцениваем,
не уведомляем, не откликаемся. Действует в matcher, hh_apply, notify_bot и
tg_outreach.
"""
from __future__ import annotations

import os
from datetime import datetime, timedelta, timezone

HARD_CAP_DAYS = 45


def max_age_days() -> int:
    try:
        days = int(os.environ.get("VACANCY_MAX_AGE_DAYS", "30"))
    except ValueError:
        days = 30
    return max(1, min(days, HARD_CAP_DAYS))


def cutoff() -> datetime:
    """Граница в UTC без часового пояса — так даты лежат в SQLite."""
    return datetime.now(timezone.utc).replace(tzinfo=None) - timedelta(days=max_age_days())


def published(v) -> datetime | None:
    """Дата публикации вакансии (ORM-объект): пост, иначе запись в базу."""
    rp = getattr(v, "raw_post", None)
    d = getattr(rp, "posted_at", None) if rp is not None else None
    d = d or getattr(v, "created_at", None)
    if d is not None and d.tzinfo is not None:
        d = d.astimezone(timezone.utc).replace(tzinfo=None)
    return d


def is_fresh(v) -> bool:
    d = published(v)
    return d is None or d >= cutoff()


def sql_condition():
    """Условие для запросов SQLAlchemy; нужен outerjoin с RawPost."""
    from sqlalchemy import func
    from jobsignal.db import RawPost, Vacancy
    return func.coalesce(RawPost.posted_at, Vacancy.created_at) >= cutoff()


# Для сырых sqlite-запросов (notify_bot): алиасы v — vacancies, r — raw_posts.
def sqlite_condition() -> str:
    return (f"COALESCE(r.posted_at, v.created_at) >= "
            f"datetime('now', '-{max_age_days()} days')")
