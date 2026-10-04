"""Показать переписку с одним собеседником — чтобы понять, как он спрашивает.

Нужен для работы с ботами-рекрутёрами. Разведка входящих показывает только
первую строку последнего сообщения, а для разбора вопросов надо видеть, как
устроен диалог целиком: один вопрос за раз или анкета, свободный текст или
кнопки, ждёт ли бот ответа прямо сейчас.

Печатает переписку как есть — это нужно, чтобы составить банк ответов. Берите
только тех собеседников, чью переписку не жалко показать: вывод пойдёт в
терминал и, скорее всего, дальше.

    cd /opt/jobsignal_local && .venv/bin/python tools/tg_dialog.py giga_recruiter_bot [сколько]
"""
from __future__ import annotations

import asyncio
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from dotenv import load_dotenv  # noqa: E402
load_dotenv("config/.env")


async def main() -> int:
    from telethon import TelegramClient

    if len(sys.argv) < 2:
        print("укажи собеседника: tools/tg_dialog.py giga_recruiter_bot")
        return 1
    who = sys.argv[1].lstrip("@")
    limit = int(sys.argv[2]) if len(sys.argv) > 2 else 30

    api_id = int(os.environ.get("TG_API_ID") or 0)
    api_hash = (os.environ.get("TG_API_HASH") or "").strip()
    name = (os.environ.get("TG_INBOX_SESSION") or "data/tg_inbox").strip()
    if not Path(f"{name}.session").exists():
        print(f"нет файла сессии {name}.session — сначала tools/tg_login.py")
        return 1

    client = TelegramClient(name, api_id, api_hash)
    await client.start()
    ent = await client.get_entity(who)
    print(f"=== @{who}: последние {limit} сообщений (сверху старые)\n")

    msgs = []
    async for m in client.iter_messages(ent, limit=limit):
        msgs.append(m)
    for m in reversed(msgs):
        side = "Я  " if m.out else "БОТ"
        when = m.date.astimezone().strftime("%d.%m %H:%M")
        text = (m.message or "").strip()
        print(f"[{when}] {side}: {text or '(без текста)'}")
        markup = getattr(m, "reply_markup", None)
        if markup is not None:
            rows = getattr(markup, "rows", []) or []
            labels = [b.text for r in rows for b in (getattr(r, "buttons", []) or [])]
            print(f"          КНОПКИ: {labels}")
        print()

    await client.disconnect()
    return 0


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
