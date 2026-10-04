"""Matcher (Этап 3) — мульти-таргет скоринг.

Для каждой основной (is_primary) вакансии в статусе NEW одним вызовом LLM
получает % соответствия под КАЖДЫЙ профиль + обоснование. Пишет строки в
match_scores, в саму вакансию кладёт лучший балл (max по профилям) и статус:
MATCHED, если лучший балл >= порога, иначе SKIPPED.

Идемпотентно: повторный прогон берёт только NEW-вакансии. Один вызов на вакансию
(а не по одному на профиль) — экономит токены.
"""
from __future__ import annotations

import logging
import os

from sqlalchemy import func, select
from sqlalchemy.orm import selectinload

from ..db import MatchScore, Vacancy, VacancyStatus, get_session_factory
from ..llm import LLMError, LLMUnavailable, complete_json
from ..tg_contacts import is_personal, mass_handles
from .base import BaseAgent

log = logging.getLogger("jobsignal")

CV_CAP = 4000  # ограничение длины резюме в промпте (контроль токенов)
# потолок на описание вакансии: требования и обязанности почти всегда в начале,
# а без предела одна многословная вакансия стоит как десяток обычных
DESC_CAP = 3000

# Какие источники оцениваем. Замер 04.10: при 391 вакансии с баллом от 80
# отправить можно было семь — автоотклик работает только через hh, остальные
# лежат в дашборде под ручное решение. Оценка вакансии стоит около $0.006, и
# разбор накопленной очереди из 3300 вакансий обошёлся бы в двадцать долларов,
# из которых девятнадцать — за те, по которым мы всё равно не откликнемся.
#
# Неоценённые НЕ теряются: статус остаётся new, полный текст сохранён, и когда
# появится канал отправки (телеграм, формы), они оценятся сами — достаточно
# расширить эту настройку. Оценивать всё, как раньше, — только явное "*":
# пустая или забытая настройка не должна молча включать расход на всё подряд.
_raw_sources = (os.environ.get("MATCH_SOURCES") or "hh").strip()
MATCH_SOURCES: set[str] | None = (
    None if _raw_sources in ("*", "") else
    {x.strip() for x in _raw_sources.split(",") if x.strip()}
)

SYSTEM_TMPL = (
    "Ты оцениваешь, насколько вакансия подходит кандидату под каждый из его "
    "карьерных профилей. Для каждого профиля поставь score 0-100 (учитывай роль, "
    "уровень/сеньорность, ключевые навыки, домен, обязанности) и ОЧЕНЬ краткое "
    "обоснование (до 10 слов) на русском. Будь честен: нерелевантной "
    "вакансии — низкий балл. Отвечай СТРОГО одним JSON без markdown:\n"
    '{"scores": [{"profile": "<имя профиля>", "score": <int 0-100>, "reason": "<кратко>"}]}\n'
    "Оцени ВСЕ профили из списка ниже.\n\n=== ПРОФИЛИ КАНДИДАТА ===\n{profiles}"
)


def _vacancy_text(v: Vacancy) -> str:
    """Текст вакансии для оценки — исходный пост, а не пересказ парсера.

    Vacancy.description это «краткое описание обязанностей (1-2 предложения)»,
    как прямо просит промпт парсера. Матчер читал именно его и оценивал
    вакансию по одной фразе: замер 21.09 показал 184 знака в среднем у hh
    против 2854 в исходном посте — пятнадцатикратная потеря. Все баллы,
    накопленные до этой правки, посчитаны по пересказу.

    Исходный текст лежит в raw_posts.text и есть у всех источников. Пересказ
    остаётся тем, чем должен быть, — подписью для интерфейса.
    """
    full = (v.raw_post.text if v.raw_post is not None else None) or ""
    body = full.strip() or (v.description or "")
    return (
        f"Роль: {v.role or '—'}\n"
        f"Компания: {v.company or '—'}\n"
        f"Локация: {v.location or '—'}\n"
        f"Зарплата: {v.salary or '—'}\n"
        f"Описание: {body[:DESC_CAP] or '—'}"
    )


class MatcherAgent(BaseAgent):
    name = "matcher"

    def run(self, limit: int | None = None) -> dict:
        Session = get_session_factory()
        cfg = self.config
        profiles = cfg.profiles
        if not profiles:
            log.warning("[matcher] нет профилей с резюме — заполни config/profiles.yaml "
                        "и cv-файлы (config/cv_*.md)")
            return {"agent": self.name, "scored": 0, "matched": 0,
                    "skipped": 0, "errors": 0, "no_profiles": True}

        names = [p["name"] for p in profiles]
        profiles_block = "\n\n".join(
            f"[{p['name']}]\n{p['cv_text'][:CV_CAP]}" for p in profiles
        )
        system = SYSTEM_TMPL.replace("{profiles}", profiles_block)

        model = cfg.settings.anthropic_model  # для anthropic; ollama/gigachat игнорят
        threshold = cfg.settings.match_threshold
        limit = limit or cfg.settings.match_batch_limit

        scored = matched = skipped = errors = 0
        with Session() as s:
            vacs = (
                s.execute(
                    select(Vacancy)
                    # Текст поста подтягиваем сразу: иначе на каждую вакансию
                    # уходил бы отдельный запрос к raw_posts.
                    .options(selectinload(Vacancy.raw_post))
                    .where(
                        Vacancy.is_primary.is_(True),
                        # NEW — ещё не оценивались; SKIPPED без балла — сбой парсинга, ретраим
                        (Vacancy.status == VacancyStatus.new)
                        | ((Vacancy.status == VacancyStatus.skipped)
                           & (~Vacancy.match_scores.any())),
                        *([Vacancy.contact_type.in_(MATCH_SOURCES)]
                          if MATCH_SOURCES else []),
                    )
                    .limit(limit)
                )
                .scalars()
                .all()
            )
            # Телеграм-вакансии с контактом-каналом отбрасываем до оценки:
            # писать туда некому, а оценка стоит денег. 24 «контакта» дают
            # 1147 вакансий из 4251 — это агрегаторы и кадровые агентства.
            # Для hh правило не действует: там отклик идёт по ссылке.
            if MATCH_SOURCES and "tg" in MATCH_SOURCES:
                mass = mass_handles(s)
                before = len(vacs)
                vacs = [v for v in vacs
                        if v.contact_type != "tg"
                        or is_personal(v.recruiter_handle, mass)]
                if before != len(vacs):
                    log.info("[matcher] пропущено телеграм-вакансий с "
                             "контактом-каналом: %d", before - len(vacs))

            # Отложенные считаем и называем вслух: иначе «оценено 0» при
            # полной очереди из телеграма выглядит как поломка, а это решение.
            deferred = 0
            if MATCH_SOURCES:
                deferred = s.execute(
                    select(func.count()).select_from(Vacancy).where(
                        Vacancy.is_primary.is_(True),
                        Vacancy.status == VacancyStatus.new,
                        Vacancy.contact_type.notin_(MATCH_SOURCES),
                    )
                ).scalar() or 0
            log.info("[matcher] к оценке вакансий: %d, профилей: %d%s",
                     len(vacs), len(names),
                     (f"; отложено {deferred} из других источников "
                      f"(оцениваем только {', '.join(sorted(MATCH_SOURCES))})"
                      if MATCH_SOURCES else ""))

            for v in vacs:
                try:
                    # cache_system=True: система (три резюме) одинакова для всех
                    # вакансий прогона — со второй вакансии она читается из кэша
                    data = complete_json(system, _vacancy_text(v), model=model,
                                         max_tokens=900, tag=f"vac#{v.id}",
                                         cache_system=True)
                except LLMUnavailable as exc:
                    # Доступа к провайдеру нет вообще — дальше по списку будет
                    # ровно то же. Оценённое сохраняем, прогон обрываем громко.
                    log.error("[matcher] ПРОГОН ПРЕРВАН: провайдер LLM "
                              "недоступен (%s). Оценено до обрыва: %d", exc, scored)
                    s.commit()
                    raise
                except LLMError as exc:
                    log.warning("[matcher] вакансия #%d: ошибка API (%s) — на ретрай", v.id, exc)
                    errors += 1
                    continue
                except ValueError as exc:
                    log.warning("[matcher] вакансия #%d: %s — помечаю SKIPPED", v.id, exc)
                    v.status = VacancyStatus.skipped
                    scored += 1
                    skipped += 1
                    continue

                by_profile = self._map_scores(data, names)
                # на случай переоценки — убираем прежние строки по этой вакансии
                for old in list(v.match_scores):
                    s.delete(old)
                best_name, best_score, best_reason = None, -1, None
                for pname in names:
                    sc = by_profile.get(pname)
                    if sc is None:
                        continue
                    score, reason = sc
                    s.merge(MatchScore(vacancy_id=v.id, profile_key=pname,
                                     score=score, reason=reason))
                    if score > best_score:
                        best_name, best_score, best_reason = pname, score, reason

                if best_score < 0:  # модель не дала ни одной валидной оценки
                    v.status = VacancyStatus.skipped
                    scored += 1
                    skipped += 1
                    continue

                # Раньше здесь лучший балл писался в v.match_score и
                # v.match_reason. Таких колонок у Vacancy нет: SQLAlchemy молча
                # клала значения в __dict__ инстанса, коммит их не сохранял, и
                # никто их не читал. Лучший балл берётся запросом
                # max(match_scores.score) — так это и делают дашборд, отправка
                # и аналитика. Порог считается по локальной переменной и был
                # верен всегда.
                v.status = VacancyStatus.matched if best_score >= threshold else VacancyStatus.skipped
                scored += 1
                if v.status == VacancyStatus.matched:
                    matched += 1
                else:
                    skipped += 1
            s.commit()

        log.info("[matcher] оценено: %d, прошли порог (%d): %d, ниже порога: %d, ошибок: %d",
                 scored, threshold, matched, skipped, errors)
        return {"agent": self.name, "scored": scored, "matched": matched,
                "skipped": skipped, "errors": errors, "deferred": deferred}

    @staticmethod
    def _map_scores(data: dict, names: list[str]) -> dict[str, tuple[int, str | None]]:
        """Сопоставляет ответ модели с известными профилями (по имени, иначе по порядку)."""
        result: dict[str, tuple[int, str | None]] = {}
        raw = data.get("scores") if isinstance(data, dict) else None
        if not isinstance(raw, list):
            return result
        lower = {n.lower(): n for n in names}
        for i, item in enumerate(raw):
            if not isinstance(item, dict):
                continue
            pname = str(item.get("profile", "")).strip()
            key = lower.get(pname.lower())
            if key is None and i < len(names):  # имя не совпало — берём по позиции
                key = names[i]
            if key is None:
                continue
            try:
                score = int(round(float(item.get("score", 0))))
            except (TypeError, ValueError):
                continue
            score = max(0, min(100, score))
            reason = item.get("reason")
            result[key] = (score, str(reason)[:500] if reason else None)
        return result
