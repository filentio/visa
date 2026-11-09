"""Сторож входящих в телеграме: не дать вопросу остаться без ответа.

Зачем. Рекрутёрские боты пишут в личку и ждут ответа: ГигаРекрутер Сбера
ведёт полноценное интервью с уточнениями по ответам, по три-четыре круга
вглубь. Пропустить такой вопрос — значит затянуть рассмотрение на дни, а
заметить его в общем потоке личных сообщений легко не сразу.

ОСТОРОЖНО С ВЫВОДАМИ. Сначала было записано, будто молчание закрывает отклик:
03.10 бот спросил в 22:41, в 22:54 пришло «статус изменён». Проверка 04.10
показала обратное — тот отклик (AI Agent & Capability Owner) ушёл в INTERVIEW.
Смена статуса не была отказом. Ценность сторожа в скорости ответа, а не в
спасении от выдуманной угрозы.

Что делает сторож. Раз в несколько минут смотрит новые входящие в личных
диалогах и шлёт уведомление, если написал:

  * бот-рекрутёр из списка TG_WATCH_BOTS;
  * любой контакт, который есть в нашей базе вакансий, — то есть рекрутёр,
    которому мы писали или чью вакансию собрали.

Ничего не отвечает сам: интервью идёт под именем человека, и выдуманная
деталь дороже сэкономленного времени. Задача сторожа — чтобы человек узнал
сразу, а не через сутки.

Чтение входящих блокировок в телеграме не вызывает: это обычное поведение,
в отличие от массовой рассылки.
"""
from __future__ import annotations

import asyncio
import html
import json
import logging
import os
from pathlib import Path

from jobsignal.db import Vacancy, get_session_factory

log = logging.getLogger("jobsignal")

# Боты-рекрутёры, чьи сообщения всегда важны. ГигаРекрутер Сбера ведёт
# интервью; hh_rabota_bot сообщает о приглашениях раньше, чем их увидит наш
# ночной обход откликов.
WATCH_BOTS = {
    h.strip().lstrip("@").lower()
    for h in (os.environ.get("TG_WATCH_BOTS")
              or "giga_recruiter_bot,hh_rabota_bot").split(",")
    if h.strip()
}

# Сколько последних диалогов просматривать. Новые сообщения поднимают диалог
# наверх, поэтому глубина нужна небольшая.
DIALOGS = int(os.environ.get("TG_WATCH_DIALOGS", "25"))

# Отметка о прочитанном: что уже показывали человеку. Файл, а не БД —
# состояние постороннее для предметной области и переживает перезапуск.
STATE = Path(os.environ.get("TG_WATCH_STATE", "data/tg_watch.json"))

PREVIEW = int(os.environ.get("TG_WATCH_PREVIEW", "600"))

# Боты, которым готовим черновик ответа. Не все подряд: hh_rabota_bot шлёт
# уведомления, отвечать там нечего, а ГигаРекрутер ведёт интервью.
DRAFT_BOTS = {
    h.strip().lstrip("@").lower()
    for h in (os.environ.get("TG_DRAFT_BOTS") or "giga_recruiter_bot").split(",")
    if h.strip()
}
# Сколько последних сообщений диалога давать для связности.
CONTEXT_MSGS = int(os.environ.get("TG_DRAFT_CONTEXT", "6"))

class TGWatchBroken(RuntimeError):
    """Сторож не может работать: нет сессии или ключей."""


def _load() -> dict:
    try:
        return json.loads(STATE.read_text())
    except Exception:  # noqa: BLE001 — нет файла или битый: начинаем заново
        return {}


def _save(data: dict) -> None:
    try:
        STATE.parent.mkdir(parents=True, exist_ok=True)
        STATE.write_text(json.dumps(data, ensure_ascii=False, indent=1))
    except Exception as exc:  # noqa: BLE001
        log.warning("[tg_watch] отметка не сохранена: %s", exc)


def _known_handles() -> set[str]:
    """Контакты из нашей базы: их сообщения — это ответы на наши отклики."""
    s = get_session_factory()()
    try:
        return {
            (h or "").lstrip("@").lower()
            for (h,) in s.query(Vacancy.recruiter_handle)
            .filter(Vacancy.recruiter_handle.isnot(None)).distinct().all()
            if h
        }
    finally:
        s.close()


def _notify(who: str, title: str, text: str, buttons: list[str]) -> None:
    from jobsignal.agents.notify_bot import _send

    body = html.escape(text[:PREVIEW]) + ("…" if len(text) > PREVIEW else "")
    lines = [f"💬 <b>{html.escape(title)}</b>", f"@{html.escape(who)}", "", body]
    if buttons:
        lines += ["", "<i>кнопки: " + html.escape(", ".join(buttons)) + "</i>"]
    lines += ["", f'<a href="https://t.me/{who}">открыть диалог</a>']
    _send("\n".join(lines))


async def _run_async() -> dict:
    try:
        from telethon import TelegramClient
    except ImportError as exc:
        raise TGWatchBroken("нет библиотеки telethon") from exc

    api_id = int(os.environ.get("TG_API_ID") or 0)
    api_hash = (os.environ.get("TG_API_HASH") or "").strip()
    session = (os.environ.get("TG_INBOX_SESSION") or "data/tg_inbox").strip()
    if not api_id or not api_hash:
        raise TGWatchBroken("TG_API_ID / TG_API_HASH не заданы")
    if not Path(f"{session}.session").exists():
        raise TGWatchBroken(f"нет файла сессии {session}.session — "
                            f"сначала tools/tg_login.py")

    seen = _load()
    known = _known_handles()
    checked = notified = drafted = 0

    client = TelegramClient(session, api_id, api_hash)
    await client.start()
    try:
        async for dialog in client.iter_dialogs(limit=DIALOGS):
            if dialog.is_channel or dialog.is_group:
                continue
            msg = dialog.message
            if msg is None or getattr(msg, "out", False):
                continue          # наше же сообщение — не повод будить
            ent = dialog.entity
            uname = (getattr(ent, "username", "") or "").lower()
            if not uname:
                continue          # без имени пользователя ссылку не дать
            is_watched = uname in WATCH_BOTS
            is_known = uname in known
            if not (is_watched or is_known):
                continue

            checked += 1
            key = str(dialog.id)
            if seen.get(key) == msg.id:
                continue          # это сообщение уже показывали

            markup = getattr(msg, "reply_markup", None)
            buttons = []
            if markup is not None:
                buttons = [b.text for r in (getattr(markup, "rows", []) or [])
                           for b in (getattr(r, "buttons", []) or [])]
            title = ("Бот-рекрутёр" if is_watched
                     else "Рекрутёр из нашей базы")
            text = getattr(msg, "message", "") or ""
            _notify(uname, title, text, buttons)
            seen[key] = msg.id
            notified += 1

            # Черновик с кнопками «Отправить / Изменить / Пропустить»
            # (jobsignal/recruiter_replies.py) — интервьюирующим ботам и живым
            # рекрутёрам из нашей базы, только на текст: к сообщению с
            # кнопками ответ печатать не нужно, там выбор.
            if (uname in DRAFT_BOTS or is_known) and not buttons and len(text) > 20:
                from jobsignal import recruiter_replies
                history = []
                async for m in client.iter_messages(ent, limit=CONTEXT_MSGS):
                    body = (m.message or "").strip()
                    if body:
                        history.append((bool(m.out), body))
                history.reverse()
                draft = recruiter_replies.make_draft(history)
                if draft:
                    recruiter_replies.offer(uname, msg.id, draft)
                    drafted += 1
    finally:
        await client.disconnect()

    _save(seen)
    log.info("[tg_watch] диалогов под наблюдением: %d, новых сообщений: %d, "
             "черновиков: %d", checked, notified, drafted)
    return {"agent": "tg_watch", "watched": checked, "notified": notified,
            "drafted": drafted}


class TGWatchAgent:
    name = "tg_watch"

    def run(self) -> dict:
        from jobsignal.tg_lock import tg_session_lock
        with tg_session_lock(wait_s=20) as got:
            if not got:
                # Сессией занят поиск каналов или сверка исходящих. Это не
                # сбой: следующий прогон через 5 минут.
                log.info("[tg_watch] сессия Telegram занята другим процессом — "
                         "прогон пропущен")
                return {"agent": "tg_watch", "watched": 0, "notified": 0,
                        "drafted": 0, "skipped": True}
            return asyncio.run(_run_async())
