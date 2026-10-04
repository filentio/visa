"""Сторож входящих в телеграме: не дать вопросу остаться без ответа.

Повод. Разбор переписки 04.10 показал, во что обходится молчание:

    [03.10 22:41] ГигаРекрутер Сбера: получил ваш отклик на AI Agent &
                  Capability Owner. Будет удобно ответить на вопросы?
    [03.10 22:54] Статус вашего рассмотрения был изменён.

Тринадцать минут — и отклик закрыт. То же 16.09. Бот Сбера не присылает
анкету, он ведёт настоящее интервью с уточнениями по ответам, и ждёт
недолго. Каждый пропущенный вопрос — выброшенный отклик, причём по самым
интересным вакансиям.

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
    checked = notified = 0

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
            _notify(uname, title, getattr(msg, "message", "") or "", buttons)
            seen[key] = msg.id
            notified += 1
    finally:
        await client.disconnect()

    _save(seen)
    log.info("[tg_watch] диалогов под наблюдением: %d, новых сообщений: %d",
             checked, notified)
    return {"agent": "tg_watch", "watched": checked, "notified": notified}


class TGWatchAgent:
    name = "tg_watch"

    def run(self) -> dict:
        return asyncio.run(_run_async())
