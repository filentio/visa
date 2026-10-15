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


_HH_ID_RE = re.compile(r"hh\.(?:ru|kz)/vacancy/(\d+)")


def signature(v: Vacancy) -> str:
    """Ключ дубля.

    Номер вакансии hh — точный ключ, и он главнее названия. Раньше ключ
    строился только из роли и компании, и пять разных «Product Manager» Сбера
    склеивались в одну группу: основной оставалась одна, четыре настоящие
    вакансии не доходили ни до оценки, ни до отклика. Репост той же вакансии
    hh из телеграма при этом по-прежнему склеивается — номер тот же.

    Вакансия без роли и без описания раньше давала ключ от пустой строки, и
    все такие сливались в одну группу. Теперь у каждой свой ключ.
    """
    m = _HH_ID_RE.search(v.link or "")
    if m:
        return f"hh{m.group(1)}"
    role = _normalize(v.role or "")
    company = _normalize(v.company or "")
    if role and company:
        base = f"{role}|{company}"
    elif role:
        base = f"{role}|{_normalize((v.description or '')[:60])}"
    else:
        base = _normalize((v.description or "")[:120])
    if not base:
        return f"id{v.id}"
    return hashlib.sha1(base.encode("utf-8")).hexdigest()[:16]


def _primary_rank(v: Vacancy):
    """Кто в группе основной. Сначала уже оценённая (matched), иначе основная
    могла смениться между прогонами, и оценённая вакансия молча выпадала из
    очереди отклика. Потом самое полное описание, потом самая ранняя — чтобы
    выбор был устойчивым."""
    return (0 if v.status == VacancyStatus.matched else 1,
            -len(v.description or ""), v.id or 0)


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
                items.sort(key=_primary_rank)
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
