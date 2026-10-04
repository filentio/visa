"""Вход в телеграм для чтения входящих — один раз, прямо на сервере.

Зачем. Сборщик читает публичные превью каналов (t.me/s/…) и личных сообщений
не видит. А нам нужны две вещи из личной переписки:

  * ответы рекрутёров на отклики — чтобы считать отдачу телеграм-канала так
    же, как считаем её для hh (28% ответов, 4 собеседования);
  * вопросы ботов-рекрутёров (Сбер и подобные) — чтобы готовить ответы.

Для этого нужен MTProto: вход как настоящее приложение, по номеру и коду.
Экран не требуется — код приходит в телеграм, вводится здесь.

ЧТО НУЖНО ЗАРАНЕЕ. Ключи разработчика с my.telegram.org (бесплатно, пять
минут): вход по номеру → API development tools → создать приложение →
api_id и api_hash. Вписать в config/.env:

    TG_API_ID=1234567
    TG_API_HASH=abcdef0123456789abcdef0123456789

ПРО КАКОЙ АККАУНТ. Файл сессии даёт полный доступ к аккаунту и будет лежать
на сервере. Чтение входящих блокировок не вызывает — это обычное поведение,
в отличие от массовой рассылки. Но если тот же аккаунт позже использовать
для отправки, риск распространится и на него. Поэтому разумнее заводить
телеграм-часть на отдельном номере: тогда худший исход — потеря второго
аккаунта, а не основного со всей личной перепиской.

    cd /opt/jobsignal_local && .venv/bin/python tools/tg_login.py
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
    try:
        from telethon import TelegramClient
    except ImportError:
        print("нет библиотеки telethon: .venv/bin/pip install telethon")
        return 1

    api_id = int(os.environ.get("TG_API_ID") or 0)
    api_hash = (os.environ.get("TG_API_HASH") or "").strip()
    if not api_id or not api_hash:
        print("TG_API_ID и TG_API_HASH не заданы в config/.env.")
        print("Получить: my.telegram.org → API development tools.")
        return 1

    # Отдельное имя сессии: сборщик каналов живёт своей жизнью, и смешивать
    # файлы не стоит — иначе перелогин одного ломает другое.
    name = (os.environ.get("TG_INBOX_SESSION") or "data/tg_inbox").strip()
    Path(name).parent.mkdir(parents=True, exist_ok=True)
    print(f"файл сессии: {name}.session\n")

    client = TelegramClient(name, api_id, api_hash)
    await client.start()          # спросит номер и код прямо здесь
    me = await client.get_me()
    print(f"\nвошли как: {me.first_name or ''} "
          f"@{me.username or '—'} (id {me.id})")
    print("Файл сессии сохранён. Второй раз вход не потребуется.")
    print("\nЭтот файл даёт полный доступ к аккаунту — в git он не попадёт")
    print("(*.session в .gitignore), но и копировать его никуда не надо.")
    await client.disconnect()
    return 0


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
