"""Занести в базу телеграм-отправки, сделанные руками до появления системы.

Зачем. По hh отдача измерена: 82 известных исхода, 23 ответа, 4 собеседования.
По телеграму в базе ноль — не потому что не писали, а потому что писали мимо
системы. Сравнить каналы нечем, и решение «автоматизировать телеграм или нет»
пришлось бы принимать на ощупь.

Факты лежат в самом телеграме: диалог с рекрутёром и в нём наше исходящее
сообщение. Сопоставляем имя пользователя с vacancies.recruiter_handle и
заводим отклик с channel='manual' и sent_at по дате первого исходящего.

Заодно чинится обратный случай: если по вакансии висит черновик, а в телеграме
видно, что отправка была, черновик становится отправленным откликом.

    cd /opt/jobsignal_local && .venv/bin/python tools/tg_backfill.py        # показать
    cd /opt/jobsignal_local && .venv/bin/python tools/tg_backfill.py --apply
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from dotenv import load_dotenv  # noqa: E402
load_dotenv("config/.env")

from sqlalchemy import select  # noqa: E402

from jobsignal.db import (Application, Vacancy, VacancyStatus,  # noqa: E402
                          get_session_factory)
from jobsignal.tg_sent import (TGSentUnavailable,  # noqa: E402
                               sent_handles_with_dates)


def main() -> int:
    apply = "--apply" in sys.argv
    try:
        sent = sent_handles_with_dates()
    except TGSentUnavailable as exc:
        print(f"телеграм недоступен: {exc}")
        return 1
    print(f"диалогов с нашим исходящим: {len(sent)}")

    s = get_session_factory()()
    created = marked = 0
    try:
        vacs = s.execute(
            select(Vacancy).where(
                Vacancy.contact_type == "tg",
                Vacancy.recruiter_handle.isnot(None),
            )
        ).scalars().all()

        for v in vacs:
            when = sent.get((v.recruiter_handle or "").lstrip("@").lower())
            if when is None:
                continue
            # Отправленный отклик уже есть — ничего не трогаем.
            if any(a.sent_at is not None for a in v.applications):
                continue
            draft = next((a for a in v.applications if a.is_draft), None)
            if draft is not None:
                # Черновик был, а в телеграме видно отправку: значит ушло.
                draft.is_draft = False
                draft.sent_at = when
                marked += 1
            else:
                s.add(Application(
                    vacancy_id=v.id,
                    message_text="(отправлено вручную до учёта в системе)",
                    channel="manual", is_draft=False, sent_at=when,
                ))
                created += 1
            if v.status in (VacancyStatus.matched, VacancyStatus.drafted):
                v.status = VacancyStatus.applied
            print(f"  #{v.id} @{v.recruiter_handle} "
                  f"{when:%d.%m.%Y} — {v.role or '—'}")

        if apply:
            s.commit()
        else:
            s.rollback()
    finally:
        s.close()

    print(f"\nновых откликов: {created}, черновиков отмечено отправленными: "
          f"{marked}")
    if not apply:
        print("это предпросмотр — повтори с --apply, чтобы записать")
    return 0


if __name__ == "__main__":
    sys.exit(main())
