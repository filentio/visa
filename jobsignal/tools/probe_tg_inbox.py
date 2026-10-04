"""Что лежит во входящих: ответы рекрутёров и вопросы ботов.

Разведка перед тем, как писать разбор. Два вопроса, на которые нужен ответ
фактами, а не предположением:

  * сколько диалогов заведено рекрутёрами, которым мы писали, — то есть есть
    ли что считать как ответы;
  * как выглядят сообщения ботов (Сбер и подобные): один вопрос за раз или
    анкета целиком, кнопки или свободный текст.

От второго напрямую зависит, можно ли вообще готовить ответы: на кнопки
отвечают нажатием, на свободный текст — сообщением, и это разный код.

Тексты сообщений НЕ печатаются целиком: в личной переписке личное. Выводим
длину, признаки вопроса и наличие кнопок.

    cd /opt/jobsignal_local && .venv/bin/python tools/probe_tg_inbox.py [сколько]
"""
from __future__ import annotations

import asyncio
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from dotenv import load_dotenv  # noqa: E402
load_dotenv("config/.env")

from jobsignal.db import Vacancy, get_session_factory  # noqa: E402


def preview(text: str, width: int = 70) -> str:
    """Первая строка сообщения — чтобы понять суть, не вываливая переписку."""
    first = (text or "").strip().splitlines()
    head = first[0] if first else ""
    return (head[:width] + "…") if len(head) > width else head


async def main() -> int:
    from telethon import TelegramClient

    limit = int(sys.argv[1]) if len(sys.argv) > 1 else 40
    api_id = int(os.environ.get("TG_API_ID") or 0)
    api_hash = (os.environ.get("TG_API_HASH") or "").strip()
    name = (os.environ.get("TG_INBOX_SESSION") or "data/tg_inbox").strip()
    if not api_id or not api_hash:
        print("TG_API_ID / TG_API_HASH не заданы — см. tools/tg_login.py")
        return 1
    if not Path(f"{name}.session").exists():
        print(f"нет файла сессии {name}.session — сначала tools/tg_login.py")
        return 1

    # Контакты, которым мы писали: с ними диалог = ответ на отклик.
    s = get_session_factory()()
    ours = {
        (h or "").lstrip("@").lower()
        for (h,) in s.query(Vacancy.recruiter_handle)
        .filter(Vacancy.recruiter_handle.isnot(None)).distinct().all()
    }
    s.close()
    print(f"контактов в базе: {len(ours)}\n")

    client = TelegramClient(name, api_id, api_hash)
    await client.start()
    ours_found = bots = people = 0

    async for dialog in client.iter_dialogs(limit=limit):
        ent = dialog.entity
        if dialog.is_channel or dialog.is_group:
            continue
        uname = (getattr(ent, "username", "") or "").lower()
        is_bot = bool(getattr(ent, "bot", False))
        mine = uname in ours
        bots += is_bot
        people += (not is_bot)
        ours_found += mine

        msg = dialog.message
        text = getattr(msg, "message", "") or ""
        buttons = bool(getattr(msg, "reply_markup", None))
        incoming = not getattr(msg, "out", False)

        mark = "БОТ " if is_bot else "чел "
        mark += "НАШ " if mine else "    "
        print(f"{mark} @{uname or '—':<24} последнее: "
              f"{'входящее' if incoming else 'наше'}, {len(text)} симв."
              f"{', КНОПКИ' if buttons else ''}")
        print(f"      {preview(text)}")

    print(f"\nдиалогов просмотрено: {limit}; ботов {bots}, людей {people}")
    print(f"из них наши контакты: {ours_found}")
    await client.disconnect()
    return 0


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
