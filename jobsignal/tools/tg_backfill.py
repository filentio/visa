"""Занести в базу телеграм-отправки, сделанные руками до появления системы.

Зачем. По hh отдача измерена: 82 известных исхода, 23 ответа, 4 собеседования.
По телеграму в базе ноль — не потому что не писали, а потому что писали мимо
системы. Сравнить каналы нечем, и решение «автоматизировать телеграм или нет»
пришлось бы принимать на ощупь.

Факты лежат в самом телеграме: диалог с рекрутёром и в нём наше исходящее
сообщение. Сопоставляем имя пользователя с vacancies.recruiter_handle и
заводим отклик с channel='backfill' и sent_at по дате первого исходящего.
Отдельный канал нужен, чтобы записи инструмента всегда можно было откатить,
не задев настоящие отметки человека (они идут как 'manual').

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
from jobsignal.tg_contacts import is_personal, mass_handles  # noqa: E402
from jobsignal.tg_sent import (TGSentUnavailable,  # noqa: E402
                               sent_handles_with_dates)


def main() -> int:
    """Записываем только то, что знаем точно.

    Первый прогон 04.10 показал, во что обходится наивное сопоставление: 176
    откликов, из которых у @TaniaR_Code 25 штук, у @vvv_hr 20, у
    @sashafedosova 18. Это агентства, которые постят поток вакансий. Одно
    наше сообщение такому контакту не значит отклик на все его вакансии, и
    запись превратила бы замер канала в выдумку — ровно то, ради чего замер
    и делался.

    Поэтому два правила. Контакты-каналы (та же проверка, что в отправке)
    пропускаем целиком. Из остальных берём только те, где у контакта ровно
    одна вакансия: тогда сообщение относится к ней однозначно. Личные
    контакты с несколькими вакансиями печатаем отдельным списком — пусть
    человек решит сам, угадывать тут нечего.
    """
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
        mass = mass_handles(s)
        vacs = s.execute(
            select(Vacancy).where(
                Vacancy.contact_type == "tg",
                Vacancy.recruiter_handle.isnot(None),
            )
        ).scalars().all()

        # Группируем по контакту: сколько у него вакансий, столько и смыслов
        # у одного отправленного сообщения.
        by_handle: dict[str, list[Vacancy]] = {}
        for v in vacs:
            h = (v.recruiter_handle or "").lstrip("@").lower()
            if h in sent:
                by_handle.setdefault(h, []).append(v)

        channels = [h for h in by_handle if not is_personal(h, mass)]
        ambiguous = {h: vs for h, vs in by_handle.items()
                     if h not in channels and len(vs) > 1}
        certain = [vs[0] for h, vs in by_handle.items()
                   if h not in channels and len(vs) == 1]
        print(f"контактов с нашим сообщением в базе вакансий: {len(by_handle)}"
              f" — из них каналов/агентств {len(channels)} (пропускаем), "
              f"с несколькими вакансиями {len(ambiguous)} (нужно решение), "
              f"однозначных {len(certain)}\n")

        for v in certain:
            when = sent[(v.recruiter_handle or "").lstrip("@").lower()]
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
                # Канал 'backfill', а НЕ 'manual'. Словом manual помечаются
                # решения человека — кнопка «✅ Отправил» в боте и кнопка в
                # дашборде. 04.10 я счёл manual признаком записей этого
                # инструмента и удалил по нему 248 строк: 176 своих и 72
                # чужих, настоящих. Вернулись из копии не все.
                #
                # Отсюда правило: инструмент, который пишет в базу пачкой,
                # помечает свои записи так, чтобы их можно было отличить от
                # всего остального и откатить, ничего не задев.
                s.add(Application(
                    vacancy_id=v.id,
                    message_text="(отправлено вручную до учёта в системе)",
                    channel="backfill", is_draft=False, sent_at=when,
                ))
                created += 1
            if v.status in (VacancyStatus.matched, VacancyStatus.drafted):
                v.status = VacancyStatus.applied
            print(f"  #{v.id} @{(v.recruiter_handle or '').lstrip('@')} "
                  f"{when:%d.%m.%Y} — {v.role or '—'}")

        if ambiguous:
            print("\nНужно решение — живой контакт, но вакансий несколько, "
                  "и к какой относится сообщение, из телеграма не видно:")
            for h, vs in sorted(ambiguous.items()):
                roles = ", ".join(f"#{v.id} {v.role or '—'}" for v in vs[:4])
                print(f"  @{h} ({len(vs)}): {roles}"
                      f"{' …' if len(vs) > 4 else ''}")

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
