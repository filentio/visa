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
from .base import BaseAgent

log = logging.getLogger("jobsignal")

CV_CAP = 4000  # ограничение длины резюме в промпте (контроль токенов)

# Базовые правила (общие для всех стилей)
BASE_SYSTEM = (
    "Ты пишешь КОРОТКОЕ сообщение рекрутёру в Telegram от первого лица (кандидат). "
    "Ровно 3-4 строки. Всегда: короткое приветствие; кто я одной фразой; интерес к "
    "конкретной роли и компании из вакансии; в конце — предложение со своей "
    "стороны, а не просьба: что готов сделать дальше (созвониться в удобное "
    "время, показать разбор похожей задачи). НЕ проси рекрутёра пересказать "
    "вакансию («расскажете подробнее о задачах?») и не предлагай прислать "
    "резюме — на hh.ru оно уходит вместе с откликом, а в телеграме прикладывается "
    "к сообщению. "
    "Возьми 2-3 ключевых слова из описания вакансии, чтобы попасть в запрос. "
    "Тон деловой и живой, без канцелярита, штампов («доброго времени суток», "
    "«заранее благодарен») и воды. НЕ выдумывай факты, цифры и места работы — "
    "только из профиля кандидата. Даты работы бери из профиля буквально: "
    "текущее место — то, у которого стоит «настоящее время», остальные "
    "описывай в прошедшем времени и не называй их «последним местом» или "
    "«текущим проектом». Без markdown, без темы письма, без подписи — "
    "только готовый текст сообщения."
)

# Стили (тональность) — пользователь выбирает в дашборде
STYLES: dict[str, dict] = {
    "metric": {
        "label": "Метрика-крючок",
        "hint": "АКЦЕНТ: сразу после представления дай ОДНУ самую сильную измеримую "
                "метрику-достижение из профиля как крючок.",
    },
    "fit": {
        "label": "Под запрос вакансии",
        "hint": "АКЦЕНТ: подчеркни точное попадание в требования — назови 2-3 совпадения "
                "опыта/навыков кандидата с описанием роли.",
    },
    "builder": {
        "label": "Хендз-он билдер",
        "hint": "АКЦЕНТ: упор на технический бэкграунд и умение самому собирать MVP на "
                "LLM/RAG/агентах (Cursor, Claude) и скорость проверки гипотез.",
    },
    "fintech": {
        "label": "Финтех / масштаб",
        "hint": "АКЦЕНТ: упор на доменный опыт в финтехе и масштаб — портфель инициатив, "
                "размер команды, бизнес-метрики (LTV, P&L).",
    },
}
DEFAULT_STYLE = "metric"


def _system_for(style: str) -> str:
    st = STYLES.get(style) or STYLES[DEFAULT_STYLE]
    return BASE_SYSTEM + "\n\n" + st["hint"]


def generate_draft(vacancy: Vacancy, profile_name: str, cv_text: str, model: str,
                   style: str = DEFAULT_STYLE) -> str:
    user = (
        f"ВАКАНСИЯ\nРоль: {vacancy.role or '—'}\n"
        f"Компания: {vacancy.company or '—'}\n"
        f"Описание: {vacancy.description or '—'}\n\n"
        f"МОЙ ПРОФИЛЬ — {profile_name}:\n{cv_text[:CV_CAP]}"
    )
    # max_tokens — предохранитель от обрыва, а не бюджет: платим за реально
    # сгенерированные токены, поэтому запас ничего не стоит. На 400 обрывов
    # ещё не было (11 писем в базе: 148-326 симв.), но кириллица идёт ~1 символ
    # на токен, и потолок был в паре десятков токенов от самого длинного
    # письма. Длину держит промпт («ровно 3-4 строки»), а не лимит —
    # поднимаем до 2000, как в notify_bot, чтобы убрать риск совсем.
    return complete_text(_system_for(style), user, model=model, max_tokens=2000,
                         tag=f"vac#{vacancy.id}", cache_system=True).strip()


def best_profile_name(v: Vacancy) -> str | None:
    best, bs = None, -1
    for m in v.match_scores:
        if m.score is not None and m.score > bs:
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
        model = cfg.settings.anthropic_model
        limit = limit or getattr(cfg.settings, "compose_batch_limit", 20)
        cvmap = {p["name"]: p["cv_text"] for p in cfg.profiles}

        made = errors = 0
        with Session() as s:
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
                pname = best_profile_name(v)
                cv = cvmap.get(pname, "")
                try:
                    text = generate_draft(v, pname or "", cv, model, style=style)
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
        log.info("[composer] черновиков создано: %d, ошибок: %d", made, errors)
        return {"agent": self.name, "drafts": made, "errors": errors}



class Composer:
    def generate(self, vacancy, style: str = "metric_hook") -> str:
        import pathlib
        best_profile = "Senior PM/PO"
        best_score = 0
        for sc in (getattr(vacancy, 'match_scores', None) or []):
            score = sc.score if hasattr(sc, 'score') else sc.get('score', 0)
            prof = sc.profile_key if hasattr(sc, 'profile_key') else sc.get('profile_key', '')
            if score > best_score:
                best_score = score
                best_profile = prof
        profile_map = {"Senior AI PM": "ai_pm", "CPO / Head of Product": "cpo", "Senior PM/PO": "pm"}
        key = profile_map.get(best_profile, "pm")
        cv_path = pathlib.Path("config") / f"cv_{key}.md"
        cv_text = cv_path.read_text() if cv_path.exists() else ""
        # Раньше здесь бралось GIGACHAT_MODEL (по умолчанию "GigaChat") — у
        # Anthropic это 404 not_found_error, и кнопка «сгенерировать письмо» в
        # дашборде падала. Письмо — качество важнее цены, поэтому основная модель.
        from ..config import get_config
        model = get_config().settings.anthropic_model
        return generate_draft(vacancy, best_profile, cv_text, model, style=style)
