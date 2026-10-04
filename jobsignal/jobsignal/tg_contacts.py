"""Отделить живых рекрутёров от каналов и агрегаторов.

Зачем. Замер 04.10 по 4251 телеграм-вакансии с контактом:

    контакт встречается | таких контактов | вакансий
    1 раз               |             914 |      914
    2 раза              |             293 |      586
    3-5                 |             162 |      596
    6-20                |             108 |     1039
    21 и больше         |              24 |     1147

Двадцать четыре «контакта» дают 1147 вакансий. Человек не публикует сто
сорок вакансий — это каналы, агрегаторы и кадровые агентства. Писать им
бесполезно, а в статистике они выглядят как доступные для отклика.

Правило простое и бесплатное: контакт, встречающийся чаще MASS_MIN раз,
личным не считаем. Никакой модели, обычный запрос к своей же базе.

Порог 6 выбран по таблице: до пяти включительно — похоже на человека,
который ведёт несколько вакансий; дальше начинается резкий скачок числа
вакансий на контакт.
"""
from __future__ import annotations

import logging
import os

from sqlalchemy import func, select

from .db import Vacancy

log = logging.getLogger("jobsignal")

MASS_MIN = int(os.environ.get("TG_MASS_HANDLE_MIN", "6"))


def mass_handles(session) -> set[str]:
    """Контакты, которые встречаются слишком часто, чтобы быть людьми."""
    rows = session.execute(
        select(Vacancy.recruiter_handle)
        .where(Vacancy.recruiter_handle.isnot(None))
        .group_by(Vacancy.recruiter_handle)
        .having(func.count() >= MASS_MIN)
    ).scalars().all()
    return {h for h in rows if h}


def is_personal(handle: str | None, mass: set[str]) -> bool:
    """Похож ли контакт на живого человека, а не на канал."""
    return bool(handle) and handle not in mass
