"""Отклик в телеграм полуавтоматом: система готовит, человек отправляет.

Зачем именно так. Телеграм — самый большой неиспользованный канал: 639 вакансий
за месяц с живым контактом рекрутёра против десяти откликов в день через hh.
Но автоматическая рассылка в личные сообщения — ровно тот образец поведения, за
который Telegram блокирует аккаунты, и потерять можно не сервис, а аккаунт со
всей перепиской.

Поэтому отправляет человек. Система делает всё остальное: отбирает вакансию,
пишет письмо, присылает его вместе со ссылкой на рекрутёра. Остаётся открыть,
вставить, отправить и нажать «Отправил» — секунды на вакансию.

Это не только осторожность, но и способ узнать, стоит ли канал усилий. Для hh
отдача измерена: 82 известных исхода, 23 ответа, 4 собеседования. Для телеграма
нет ничего. Три десятка ручных отправок дадут число, с которым можно сравнивать
— и тогда решение про автоматизацию будет опираться на замер, а не на догадку.

Кого НЕ берём: контакты-каналы и агрегаторы (см. tg_contacts), вакансии без
оценки и те, по которым отклик уже есть.
"""
from __future__ import annotations

import html
import logging
import os

from sqlalchemy import func, select

from jobsignal.db import (Application, MatchScore, Vacancy, VacancyStatus,
                          get_session_factory, utcnow)
from jobsignal.tg_contacts import is_personal, mass_handles

log = logging.getLogger("jobsignal")

# Сколько вакансий предлагать за прогон. Важнее не производительность, а чтобы
# человек реально разобрал пачку: двадцать карточек подряд никто не осилит.
BATCH = int(os.environ.get("TG_OUTREACH_BATCH", "5"))
# Порог балла. Отдельный от hh: там за отклик платит только квота, здесь —
# внимание человека, поэтому планка выше.
THRESHOLD = int(os.environ.get("TG_OUTREACH_THRESHOLD", "80"))


class TGOutreachAgent:
    name = "tg_outreach"

    def run(self, limit: int | None = None) -> dict:
        from jobsignal.agents.composer import Composer
        from jobsignal.agents.notify_bot import _send

        batch = limit or BATCH
        sf = get_session_factory()
        s = sf()
        try:
            mass = mass_handles(s)
            best = (
                select(MatchScore.vacancy_id,
                       func.max(MatchScore.score).label("score"))
                .group_by(MatchScore.vacancy_id).subquery()
            )
            rows = s.execute(
                select(Vacancy, best.c.score)
                .join(best, best.c.vacancy_id == Vacancy.id)
                .where(
                    Vacancy.contact_type == "tg",
                    Vacancy.is_primary.is_(True),
                    Vacancy.recruiter_handle.isnot(None),
                    Vacancy.status.in_([VacancyStatus.matched,
                                        VacancyStatus.drafted]),
                    best.c.score >= THRESHOLD,
                    ~Vacancy.applications.any(Application.sent_at.isnot(None)),
                )
                .order_by(best.c.score.desc(), Vacancy.created_at.desc())
            ).all()

            queue = [(v, sc) for v, sc in rows
                     if is_personal(v.recruiter_handle, mass)][:batch]
            log.info("[tg_outreach] очередь по порогу %d%%: %d вакансий "
                     "(показываю %d)", THRESHOLD, len(rows), len(queue))
            if not queue:
                return {"agent": self.name, "queue": 0, "offered": 0}

            composer = Composer()
            offered = failed = 0
            for v, score in queue:
                # Письмо пишем здесь, а не заранее: человек может пачку и не
                # разобрать, а вызов модели стоит денег.
                try:
                    letter = composer.generate(v)
                except Exception as exc:  # noqa: BLE001 — одна вакансия не
                    # должна рушить пачку
                    log.warning("[tg_outreach] #%d: письмо не составлено: %s",
                                v.id, exc)
                    failed += 1
                    continue

                handle = (v.recruiter_handle or "").lstrip("@")
                head = (f"📨 <b>Отклик в телеграм</b> · {score}%\n"
                        f"{html.escape(v.role or 'вакансия')}"
                        f"{' · ' + html.escape(v.company) if v.company else ''}\n"
                        f"@{html.escape(handle)}")
                body = f"<pre>{html.escape(letter)}</pre>"
                tail = (f'<a href="https://t.me/{handle}">открыть диалог</a> — '
                        f"скопируй письмо, отправь и отметь ниже")
                markup = {"inline_keyboard": [[
                    {"text": "✅ Отправил", "callback_data": f"applied:{v.id}"},
                    {"text": "⏭ Пропустить", "callback_data": f"tgskip:{v.id}"},
                ]]}
                if _send(f"{head}\n\n{body}\n\n{tail}", markup) is None:
                    failed += 1
                    continue

                # Черновик в базе: так письмо не потеряется и видно, что по
                # вакансии уже предлагали отклик. sent_at пуст — квоту не
                # занимает, отправки ещё не было.
                if not any(a.is_draft for a in v.applications):
                    s.add(Application(vacancy_id=v.id, message_text=letter,
                                      channel="telegram", is_draft=True,
                                      created_at=utcnow()))
                v.draft_text = letter
                v.status = VacancyStatus.drafted
                offered += 1
            s.commit()
        finally:
            s.close()

        log.info("[tg_outreach] предложено: %d, не вышло: %d", offered, failed)
        return {"agent": self.name, "queue": len(queue), "offered": offered,
                "failed": failed}
