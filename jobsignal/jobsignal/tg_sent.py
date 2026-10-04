"""Кому в телеграме уже писали — по самому телеграму, а не по нашей базе.

Зачем. Первая же пачка полуавтоматической отправки выдала пять карточек, и по
всем пяти человек уже писал руками. Канал в таком виде бесполезен: разбирать
предложения, половина которых — повтор, никто не станет.

Эвристики тут не нужны, факт лежит рядом. Мы и так входим в телеграм по
MTProto ради сторожа входящих, а значит видим переписку. Если в диалоге с
контактом есть ИСХОДЯЩЕЕ сообщение — писали. Это не вывод, это наблюдение, и
оно верно независимо от того, прошла отправка через систему или руками, до
появления системы или после.

Стоимость: один проход по списку диалогов (их сотни, не тысячи) с чтением
одного последнего исходящего в каждом. Секунды, лимитов телеграма не трогает.
"""
from __future__ import annotations

import asyncio
import logging
import os
from pathlib import Path

log = logging.getLogger("jobsignal")

# Сколько диалогов просматривать. В отличие от сторожа входящих, где важны
# только свежие, здесь нужна вся история: человек мог написать рекрутёру
# месяц назад, и диалог давно утонул.
DIALOGS = int(os.environ.get("TG_SENT_DIALOGS", "500"))


class TGSentUnavailable(RuntimeError):
    """Телеграм недоступен — список исходящих получить неоткуда."""


async def _collect(with_dates: bool) -> dict[str, object]:
    try:
        from telethon import TelegramClient
    except ImportError as exc:
        raise TGSentUnavailable("нет библиотеки telethon") from exc

    api_id = int(os.environ.get("TG_API_ID") or 0)
    api_hash = (os.environ.get("TG_API_HASH") or "").strip()
    name = (os.environ.get("TG_INBOX_SESSION") or "data/tg_inbox").strip()
    if not api_id or not api_hash:
        raise TGSentUnavailable("TG_API_ID / TG_API_HASH не заданы")
    if not Path(f"{name}.session").exists():
        raise TGSentUnavailable(f"нет файла сессии {name}.session — "
                                f"сначала tools/tg_login.py")

    found: dict[str, object] = {}
    client = TelegramClient(name, api_id, api_hash)
    await client.start()
    try:
        async for dialog in client.iter_dialogs(limit=DIALOGS):
            if dialog.is_channel or dialog.is_group:
                continue
            uname = (getattr(dialog.entity, "username", "") or "").lower()
            if not uname:
                continue
            # Быстрый путь: последнее сообщение диалога наше — писали, и
            # перебирать историю незачем. Так закрывается большинство.
            msg = dialog.message
            if msg is not None and getattr(msg, "out", False):
                found[uname] = msg.date if with_dates else True
                continue
            # Иначе ищем самое раннее наше сообщение: для задним числом
            # проставленных откликов важна дата первой отправки, а не
            # последней — именно она начало разговора.
            first = None
            async for m in client.iter_messages(dialog.entity, limit=200,
                                                reverse=True):
                if getattr(m, "out", False):
                    first = m
                    break
            if first is not None:
                found[uname] = first.date if with_dates else True
    finally:
        await client.disconnect()
    return found


def sent_handles() -> set[str]:
    """Хэндлы, которым мы когда-либо писали. Бросает TGSentUnavailable."""
    return set(asyncio.run(_collect(False)))


def sent_handles_with_dates() -> dict[str, object]:
    """То же плюс дата первого исходящего — для простановки задним числом."""
    return asyncio.run(_collect(True))
