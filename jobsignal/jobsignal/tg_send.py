"""Отправка отклика рекрутёру в Telegram с аккаунта кандидата — по кнопке.

Зачем. До 09.11 отклик в телеграм шёл руками: бот присылал письмо и PDF
(«перешли рекрутёру»), человек открывал чат, копировал, отправлял и жал
«Откликнулся». Теперь под письмом кнопка «📤 Отправить»: нажатие — это
разрешение, и система сама отправляет письмо и резюме с аккаунта кандидата
(та же сессия Telethon, что у сторожа входящих, data/tg_inbox.session).

Защиты:
  * только по нажатию человека — автоматической рассылки здесь нет;
  * если в диалоге с этим контактом уже есть наше исходящее сообщение
    (писали руками или раньше системой) — не отправляем, повтор хуже тишины;
  * не больше TG_SEND_PER_DAY (10) отправок в сутки: телеграм ограничивает
    аккаунты, которые много пишут незнакомым людям;
  * сессию берём под общим замком (jobsignal/tg_lock.py).
"""
from __future__ import annotations

import asyncio
import logging
import os
import shutil
import tempfile
from pathlib import Path

log = logging.getLogger("jobsignal")


class TGSendError(RuntimeError):
    """Отправить не получилось — текст для человека в сообщении."""


class AlreadyContacted(TGSendError):
    """В диалоге уже есть наше исходящее — повтор не отправляем."""


def _resume_name(pdf: Path) -> str:
    name = (os.environ.get("EMAIL_FROM_NAME") or os.environ.get("CANDIDATE_NAME") or "").strip()
    return f"Резюме {name}{pdf.suffix}" if name else pdf.name


async def _send(handle: str, text: str, pdf: Path | None) -> None:
    from telethon import TelegramClient

    api_id = int(os.environ.get("TG_API_ID") or 0)
    api_hash = (os.environ.get("TG_API_HASH") or "").strip()
    session = (os.environ.get("TG_INBOX_SESSION") or "data/tg_inbox").strip()
    if not api_id or not api_hash:
        raise TGSendError("на сервере не заданы TG_API_ID / TG_API_HASH")
    if not Path(f"{session}.session").exists():
        raise TGSendError("нет сессии Telegram на сервере (tools/tg_login.py)")
    client = TelegramClient(session, api_id, api_hash)
    await client.connect()
    try:
        if not await client.is_user_authorized():
            raise TGSendError("сессия Telegram на сервере не авторизована")
        try:
            entity = await client.get_entity(handle)
        except (ValueError, TypeError) as exc:
            raise TGSendError(f"контакт {handle} в Telegram не найден") from exc
        mine = await client.get_messages(entity, limit=1, from_user="me")
        if mine:
            raise AlreadyContacted(f"вы уже писали {handle} — повтор не отправляю")
        await client.send_message(entity, text)
        if pdf is not None and pdf.exists():
            # Имя файла — «Резюме Имя Фамилия.pdf», а не служебное po.pdf.
            with tempfile.TemporaryDirectory() as tmp:
                named = Path(tmp) / _resume_name(pdf)
                shutil.copyfile(pdf, named)
                await client.send_file(entity, str(named), force_document=True)
    finally:
        await client.disconnect()


def sent_today() -> int:
    import sqlite3
    try:
        c = sqlite3.connect("data/jobsignal.db")
        n = c.execute("SELECT COUNT(*) FROM applications WHERE channel = 'telegram' "
                      "AND sent_at IS NOT NULL AND sent_at >= datetime('now', '-1 day')"
                      ).fetchone()[0]
        c.close()
        return int(n or 0)
    except Exception:  # noqa: BLE001
        return 0


def send_application(handle: str, text: str, pdf: Path | None) -> None:
    """Отправить письмо и резюме. Исключение TGSendError — с причиной."""
    from jobsignal.tg_lock import tg_session_lock

    handle = "@" + handle.lstrip("@").strip()
    if not text.strip():
        raise TGSendError("пустое письмо")
    limit = int(os.environ.get("TG_SEND_PER_DAY", "10"))
    if sent_today() >= limit:
        raise TGSendError(f"сегодня уже {limit} отправок в Telegram — лимит "
                          "(TG_SEND_PER_DAY), чтобы телеграм не счёл аккаунт спамером")
    with tg_session_lock(wait_s=90) as got:
        if not got:
            raise TGSendError("сессия Telegram занята другой задачей — нажмите ещё раз через минуту")
        asyncio.run(_send(handle, text, pdf))
    log.info("[tg_send] отклик отправлен %s", handle)
