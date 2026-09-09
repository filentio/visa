"""Deduplicator (Этап 2).

Одна и та же вакансия часто висит в нескольких каналах. Группируем активные
вакансии по нормализованной сигнатуре (роль+компания, иначе — по тексту),
в каждой группе оставляем одну is_primary (с самым полным описанием),
остальные помечаем is_primary=False — дашборд покажет только основные.

Без тяжёлых эмбеддингов: нормализованный ключ ловит точные/почти точные репосты.
Фаззи-дедуп по эмбеддингам — опциональный апгрейд (sentence-transformers).
"""
from __future__ import annotations

import hashlib
import logging
import re

from sqlalchemy import select

from ..db import Vacancy, VacancyStatus, get_session_factory
from .base import BaseAgent

log = logging.getLogger("jobsignal")

# дедупим только вакансии, по которым ещё не действовали
ACTIVE = (VacancyStatus.new, VacancyStatus.matched)


def _normalize(text: str) -> str:
    text = text.lower()
    text = re.sub(r"[^\w\s]", " ", text, flags=re.UNICODE)  # убрать пунктуацию/эмодзи
    return re.sub(r"\s+", " ", text).strip()


def signature(v: Vacancy) -> str:
    """Ключ дубля: роль+компания, иначе — начало описания/текста."""
    role = _normalize(v.role or "")
    company = _normalize(v.company or "")
    if role and company:
        base = f"{role}|{company}"
    elif role:
        base = f"{role}|{_normalize((v.description or '')[:60])}"
    else:
        base = _normalize((v.description or "")[:120])
    return hashlib.sha1(base.encode("utf-8")).hexdigest()[:16]


class DeduplicatorAgent(BaseAgent):
    name = "deduplicator"

    def run(self) -> dict:
        Session = get_session_factory()
        with Session() as s:
            vacancies = (
                s.execute(select(Vacancy).where(Vacancy.status.in_(ACTIVE)))
                .scalars()
                .all()
            )
            groups: dict[str, list[Vacancy]] = {}
            for v in vacancies:
                groups.setdefault(signature(v), []).append(v)

            dup_groups = 0
            dup_count = 0
            for sig, items in groups.items():
                # основная = с самым длинным описанием (самая полная версия)
                items.sort(key=lambda x: len(x.description or ""), reverse=True)
                for i, v in enumerate(items):
                    v.dedup_group = sig
                    v.is_primary = (i == 0)
                if len(items) > 1:
                    dup_groups += 1
                    dup_count += len(items) - 1
            s.commit()

        log.info("[deduplicator] групп с дублями: %d, скрыто дублей: %d (из %d вакансий)",
                 dup_groups, dup_count, len(vacancies))
        return {"agent": self.name, "duplicate_groups": dup_groups,
                "duplicates_hidden": dup_count, "total": len(vacancies)}
