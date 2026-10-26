"""Composer (Этап 5).

Генерит КОРОТКИЙ черновик сообщения рекрутёру под вакансию, опираясь на профиль,
по которому вакансия прошла лучше всего. Semi-auto: черновик кладётся в
applications(is_draft=True), статус вакансии -> DRAFTED. Отправляет потом сам
пользователь кнопкой в дашборде.

Основной путь — генерация по кнопке в дашборде (экономит токены). ComposerAgent.run()
— батч-вариант для CLI, по умолчанию ограничен compose_batch_limit.
"""
from __future__ import annotations

import logging

from sqlalchemy import select

from ..db import Application, MatchScore, Vacancy, VacancyStatus, get_session_factory, utcnow
from ..llm import LLMError, complete_text
from ..tg_contacts import is_personal, mass_handles
from .base import BaseAgent

log = logging.getLogger("jobsignal")

# Правила и построение письма — в jobsignal/letters.py, общем для всего
# проекта. Здесь был свой промпт («ровно 3-4 строки») и свои стили, и письма из
# карточек Telegram отличались от писем кабинета и автоотклика (26.10).
from jobsignal.letters import ACCENTS as STYLES  # noqa: E402 — имя для старых вызовов
DEFAULT_STYLE = "auto"


def generate_draft(vacancy: Vacancy, profile_name: str, cv_text: str = "", model: str = "",
                   style: str = DEFAULT_STYLE, channel: str = "telegram") -> str:
    """Черновик через единый сервис писем. cv_text и model больше не нужны:
    резюме и модель сервис берёт сам. Пустое письмо — исключение, чтобы
    вызывающий не сохранил и не отправил пустоту."""
    from jobsignal import letters
    raw = getattr(vacancy, "raw_post", None)
    body = ((raw.text if raw is not None else "") or vacancy.description or "")
    letter = letters.generate(vacancy.role or "", vacancy.company or "", body,
                              profile=profile_name or None, channel=channel,
                              accent=style)
    if not letter.text:
        raise LLMError("; ".join(letter.issues) or "письмо не получено")
    return letter.text


def best_profile_name(v: Vacancy) -> str | None:
    """Лучший профиль среди действующих (config/profiles.yaml)."""
    from jobsignal import profiles
    active = set(profiles.names())
    best, bs = None, -1
    for m in v.match_scores:
        if m.profile_key in active and m.score is not None and m.score > bs:
            # В модели поле называется profile_key (в нём лежит имя профиля,
            # например «Senior AI PM») — m.profile не существовало, и
            # best_profile_name падал с AttributeError на любой вакансии,
            # у которой есть оценки.
            best, bs = m.profile_key, m.score
    return best


class ComposerAgent(BaseAgent):
    name = "composer"

    def run(self, limit: int | None = None, style: str = DEFAULT_STYLE) -> dict:
        Session = get_session_factory()
        cfg = self.config
        limit = limit or getattr(cfg.settings, "compose_batch_limit", 20)

        made = errors = skipped_mass = 0
        with Session() as s:
            # Каналы и агрегаторы из очереди убираем: письмо «здравствуйте,
            # меня заинтересовала ваша вакансия» в канал на сто сорок вакансий
            # уходит в пустоту, а в отчётах выглядит как отправленный отклик.
            mass = mass_handles(s)
            vacs = (
                s.execute(
                    select(Vacancy)
                    .where(Vacancy.status == VacancyStatus.matched,
                           Vacancy.recruiter_handle.is_not(None),
                           Vacancy.is_primary.is_(True))
                    .limit(limit)
                )
                .scalars()
                .all()
            )
            for v in vacs:
                if any(a.is_draft for a in v.applications):
                    continue  # черновик уже есть
                if not is_personal(v.recruiter_handle, mass):
                    skipped_mass += 1
                    continue
                pname = best_profile_name(v)
                try:
                    text = generate_draft(v, pname or "", style=style)
                except LLMError as exc:
                    log.warning("[composer] вакансия #%d: %s", v.id, exc)
                    errors += 1
                    continue
                # Канал обязателен: черновик пишется под TG-рекрутёра
                # (в выборке выше recruiter_handle не пуст), отправлять его
                # будет пользователь из телеграма. sent_at пуст — черновик
                # ещё никуда не ушёл и квоту не занимает.
                s.add(Application(vacancy_id=v.id, message_text=text,
                                  channel="telegram", is_draft=True))
                v.status = VacancyStatus.drafted
                made += 1
            s.commit()
        log.info("[composer] черновиков создано: %d, ошибок: %d, "
                 "пропущено каналов и агрегаторов: %d", made, errors, skipped_mass)
        return {"agent": self.name, "drafts": made, "errors": errors,
                "skipped_mass": skipped_mass}



class Composer:
    """Письмо под вакансию для Telegram (карточки отклика, run.py compose)."""

    def generate(self, vacancy, style: str = DEFAULT_STYLE) -> str:
        profile = best_profile_name(vacancy)
        return generate_draft(vacancy, profile or "", style=style, channel="telegram")
