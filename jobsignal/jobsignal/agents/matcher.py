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

from sqlalchemy import select

from ..db import MatchScore, Vacancy, VacancyStatus, get_session_factory
from ..llm import LLMError, complete_json
from .base import BaseAgent

log = logging.getLogger("jobsignal")

CV_CAP = 4000  # ограничение длины резюме в промпте (контроль токенов)

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
    return (
        f"Роль: {v.role or '—'}\n"
        f"Компания: {v.company or '—'}\n"
        f"Локация: {v.location or '—'}\n"
        f"Зарплата: {v.salary or '—'}\n"
        f"Описание: {v.description or '—'}"
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
                    .where(
                        Vacancy.is_primary.is_(True),
                        # NEW — ещё не оценивались; SKIPPED без балла — сбой парсинга, ретраим
                        (Vacancy.status == VacancyStatus.new)
                        | ((Vacancy.status == VacancyStatus.skipped)
                           & (~Vacancy.match_scores.any())),
                    )
                    .limit(limit)
                )
                .scalars()
                .all()
            )
            log.info("[matcher] к оценке вакансий: %d, профилей: %d", len(vacs), len(names))

            for v in vacs:
                try:
                    data = complete_json(system, _vacancy_text(v), model=model, max_tokens=900)
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

                v.match_score = best_score
                v.match_reason = f"{best_name}: {best_reason}" if best_reason else best_name
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
                "skipped": skipped, "errors": errors}

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
