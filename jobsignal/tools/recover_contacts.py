"""Вернуть способ отклика вакансиям, которые считались «без контакта».

Зачем. Замер 07.10 по 567 телеграм-вакансиям без recruiter_handle:

    почта рекрутёра в тексте           302
    ссылка на hh / форму / сайт         60
    другой телеграм-контакт             74
    бот отклика                          7
    ничего, кроме ссылки на свой канал 124

То есть глухих по-настоящему — пятая часть. Остальные четыреста с лишним
вакансий месяцами выбрасывались, хотя способ отклика лежал в тексте поста.
Парсер искал только телеграм-контакт рекрутёра, а почту и внешние ссылки за
контакт не считал.

Что делаем, в порядке надёжности:

  * ссылка на вакансию hh → link и contact_type='hh'. Дальше её подберёт
    обычный конвейер: оценка и автоотклик, как любую вакансию с hh;
  * телеграм-контакт, отличный от самого канала и не бот → recruiter_handle;
  * почта → recruiter_email и contact_type='email'.

Статус вакансии не трогаем: она остаётся new и попадёт в оценку сама. Канал
отправки по почте — отдельное решение, здесь только разметка.

Ссылку канала на самого себя («t.me/toplevel_job» в посте из @toplevel_job)
контактом не считаем: это подпись, а не человек.

    cd /opt/jobsignal_local && .venv/bin/python tools/recover_contacts.py
    cd /opt/jobsignal_local && .venv/bin/python tools/recover_contacts.py --apply
"""
from __future__ import annotations

import re
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from dotenv import load_dotenv  # noqa: E402
load_dotenv("config/.env")

from sqlalchemy import select  # noqa: E402
from sqlalchemy.orm import selectinload  # noqa: E402

from jobsignal.db import (Channel, RawPost, Vacancy,  # noqa: E402
                          get_session_factory)
from jobsignal.tg_contacts import mass_handles  # noqa: E402

EMAIL = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,}")
# Схема необязательна: в постах ссылки часто без неё — «hh.ru/vacancy/123».
# Первый предпросмотр с обязательным https:// нашёл ноль при шестидесяти
# по замеру.
HH = re.compile(r"(?:https?://)?(?:[a-z]+\.)?hh\.(?:ru|kz)/vacancy/\d+", re.I)
# Собака — телеграм-имя, только если перед ней не стоит часть адреса. Иначе
# «ivan@sberbank.ru» даёт контакт «@sberbank»: список почтовых доменов спасает
# от gmail и yandex, но не от корпоративной почты.
TME = re.compile(r"(?:t\.me/|(?<![A-Za-z0-9._%+-])@)([A-Za-z][A-Za-z0-9_]{4,31})\b")

# Адреса, которые не ведут к рекрутёру: служебные ящики площадок и
# заглушки из шаблонов постов.
JUNK_EMAIL = ("noreply", "no-reply", "example.", "support@", "info@hh.")


def _channel_handle(post: RawPost | None) -> str:
    ch = getattr(post, "channel", None) if post is not None else None
    return ((getattr(ch, "handle", "") or "").lstrip("@").lower())


# Признаки канала, а не человека, в самом имени. Нужны потому, что посты
# рекламируют соседние каналы: первый предпросмотр 07.10 выдал «контактами
# рекрутёров» @qa_jobs, @job_react, @devs_it и даже @workayte — один из наших
# же отслеживаемых каналов.
CHANNELISH = ("job", "vacanc", "vakans", "career", "devs_", "_it", "remote",
              "freelance", "work")


def classify(v: Vacancy, not_people: set[str]) -> tuple[str, str]:
    """(вид, значение) — лучший найденный способ отклика или ("", "").

    Порядок — по надёжности, и он важен. Почта стоит ВЫШЕ телеграма: в первой
    версии было наоборот, и пост с рекламой соседнего канала и почтой
    рекрутёра уходил в «телеграм» — ссылка на канал перебивала живой адрес.
    """
    text = (v.raw_post.text if v.raw_post is not None else "") or ""
    own = _channel_handle(v.raw_post)

    m = HH.search(text)
    if m:
        url = m.group(0)
        return "hh", url if url.lower().startswith("http") else "https://" + url

    for e in EMAIL.findall(text):
        el = e.lower()
        if any(j in el for j in JUNK_EMAIL):
            continue
        return "email", el

    for h in TME.findall(text):
        hl = h.lower()
        if hl == own or hl.endswith("bot") or hl in not_people:
            continue
        if any(p in hl for p in CHANNELISH):
            continue
        return "tg", "@" + h

    return "", ""


def main() -> int:
    apply = "--apply" in sys.argv
    s = get_session_factory()()
    kinds: Counter[str] = Counter()
    examples: dict[str, list[str]] = {}
    try:
        vacs = s.execute(
            select(Vacancy)
            .options(selectinload(Vacancy.raw_post).selectinload(RawPost.channel))
            .where(Vacancy.contact_type == "tg",
                   Vacancy.recruiter_handle.is_(None),
                   Vacancy.is_primary.is_(True))
        ).scalars().all()

        # Кто точно не человек: наши отслеживаемые каналы и контакты, которые
        # встречаются слишком часто (та же проверка, что в отправке).
        not_people = {(h or "").lstrip("@").lower()
                      for (h,) in s.query(Channel.handle).all() if h}
        not_people |= {(h or "").lstrip("@").lower() for h in mass_handles(s)}

        for v in vacs:
            kind, value = classify(v, not_people)
            kinds[kind or "нет"] += 1
            if kind and len(examples.setdefault(kind, [])) < 5:
                examples[kind].append(f"#{v.id} {(v.role or '?')[:40]} → {value}")
            if not apply or not kind:
                continue
            if kind == "hh":
                v.link = value
                v.contact_type = "hh"
            elif kind == "tg":
                v.recruiter_handle = value
            elif kind == "email":
                v.recruiter_email = value
                v.contact_type = "email"

        if apply:
            s.commit()
        else:
            s.rollback()
    finally:
        s.close()

    print(f"телеграм-вакансий без контакта: {sum(kinds.values())}")
    names = {"hh": "ссылка на вакансию hh → в автоотклик",
             "tg": "телеграм-контакт рекрутёра",
             "email": "почта рекрутёра",
             "нет": "способа отклика не нашлось"}
    for k in ("hh", "tg", "email", "нет"):
        if kinds.get(k):
            print(f"  {names[k]}: {kinds[k]}")
            for ex in examples.get(k, []):
                print(f"      {ex}")
    if not apply:
        print("\nэто предпросмотр — повтори с --apply, чтобы записать")
    else:
        print("\nзаписано. Вакансии остались в статусе new и попадут в оценку "
              "на ближайшем прогоне.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
