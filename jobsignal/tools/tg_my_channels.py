"""Найти в своих подписках каналы с вакансиями и предложить к отслеживанию.

Зачем. Список из 35 каналов собран вручную и с тех пор не пополнялся, а
подписки живут своей жизнью: человек находит паблик, подписывается и забывает.
Эти каналы уже отобраны им самим — это лучший источник кандидатов, чем поиск
по tgstat, и он не требует ни денег, ни внешних сервисов.

Как решаем, что канал про вакансии. Не по названию: «Perforum_jobs» угадать
легко, а «careerspace» или «bankman_20» — нет, и наоборот, «AI Jobs Digest»
может оказаться рекламной рассылкой. Поэтому берём по пять последних
сообщений и показываем их модели: публикуются ли здесь вакансии регулярно.
Судим по содержимому, а не по вывеске.

Найденное кладём в channel_candidates со статусом pending и печатаем. Это
предложение, а не решение: добавить в отслеживание — отдельный шаг с --add,
потому что лишний канал стоит денег на каждом разборе.

    cd /opt/jobsignal_local && .venv/bin/python tools/tg_my_channels.py
    cd /opt/jobsignal_local && .venv/bin/python tools/tg_my_channels.py \
        --add --only product_jobs,products_jobs_projects
"""
from __future__ import annotations

import asyncio
import json
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from dotenv import load_dotenv  # noqa: E402
load_dotenv("config/.env")

from jobsignal.db import (Channel, ChannelCandidate,  # noqa: E402
                          get_session_factory, utcnow)

# Сколько последних сообщений показывать модели. Пяти хватает, чтобы отличить
# поток вакансий от новостей и рекламы, и это дёшево.
SAMPLE = int(os.environ.get("TG_SAMPLE_MSGS", "5"))
# Сколько знаков от каждого сообщения. Вакансия узнаётся по началу.
SAMPLE_CHARS = 400

SYSTEM = (
    "Ты определяешь, публикует ли телеграм-канал вакансии. Тебе дают "
    "название канала и несколько последних сообщений.\n\n"
    "Канал подходит, если вакансии публикуются регулярно и это его основное "
    "содержание. НЕ подходят: новости отрасли, обучение и курсы, реклама, "
    "личные блоги, каналы с вакансиями от случая к случаю, чаты общения.\n\n"
    "Для каждого канала верни решение и одной фразой причину, а для "
    "подходящих — нишу двумя-тремя словами (например «продакт-менеджмент», "
    "«аналитика данных», «финансы и C-level»).\n\n"
    "Отвечай СТРОГО одним JSON без markdown:\n"
    '{"channels": [{"handle": "<как дали>", "fits": true|false, '
    '"niche": "<ниша или пусто>", "reason": "<кратко>"}]}'
)


async def collect() -> list[dict]:
    from telethon import TelegramClient

    api_id = int(os.environ.get("TG_API_ID") or 0)
    api_hash = (os.environ.get("TG_API_HASH") or "").strip()
    name = (os.environ.get("TG_INBOX_SESSION") or "data/tg_inbox").strip()
    if not Path(f"{name}.session").exists():
        print(f"нет файла сессии {name}.session — сначала tools/tg_login.py")
        return []

    s = get_session_factory()()
    try:
        known = {(h or "").lstrip("@").lower()
                 for (h,) in s.query(Channel.handle).all() if h}
        seen_cand = {(h or "").lstrip("@").lower()
                     for (h,) in s.query(ChannelCandidate.handle).all() if h}
    finally:
        s.close()

    out: list[dict] = []
    client = TelegramClient(name, api_id, api_hash)
    await client.start()
    try:
        async for dialog in client.iter_dialogs():
            ent = dialog.entity
            # Только каналы и супергруппы: личные диалоги это переписка, а не
            # источник вакансий, и читать их здесь незачем.
            if not (dialog.is_channel or dialog.is_group):
                continue
            uname = (getattr(ent, "username", "") or "").lower()
            if not uname:
                continue      # без имени канал не подписать на сбор
            if uname in known:
                continue      # уже отслеживаем
            msgs = []
            async for m in client.iter_messages(ent, limit=SAMPLE):
                body = (m.message or "").strip()
                if body:
                    msgs.append(body[:SAMPLE_CHARS])
            if not msgs:
                continue      # пустой или только медиа — судить не о чем
            out.append({
                "handle": uname,
                "title": getattr(ent, "title", "") or uname,
                "subs": getattr(ent, "participants_count", 0) or 0,
                "msgs": msgs,
                "known_candidate": uname in seen_cand,
            })
    finally:
        await client.disconnect()
    return out


def judge(items: list[dict]) -> dict[str, dict]:
    """Решение модели по каждому каналу. Пачками, чтобы не платить за вызов
    на канал: двести подписок это двести запросов вместо десяти."""
    from jobsignal.config import get_config
    from jobsignal.llm import complete_json

    model = get_config().settings.anthropic_model
    verdicts: dict[str, dict] = {}
    BATCH = 10
    for i in range(0, len(items), BATCH):
        chunk = items[i:i + BATCH]
        blocks = []
        for it in chunk:
            sample = "\n---\n".join(it["msgs"])
            blocks.append(f"[{it['handle']}] {it['title']}\n{sample}")
        user = "\n\n=====\n\n".join(blocks)
        try:
            data = complete_json(SYSTEM, user, model=model, max_tokens=1200,
                                 tag="tg_channels", cache_system=True)
        except Exception as exc:  # noqa: BLE001 — одна пачка не рушит обход
            print(f"  (пачка {i // BATCH + 1}: решение не получено — {exc})")
            continue
        for row in (data.get("channels") or []):
            h = str(row.get("handle", "")).lstrip("@").lower()
            if h:
                verdicts[h] = row
    return verdicts


def _only() -> set[str]:
    """--only a,b,c — взять лишь названные каналы.

    Нужно потому, что «публикует вакансии» и «публикует НУЖНЫЕ вакансии» —
    разные вещи. Первый обход 07.10 нашёл одиннадцать годных каналов, из
    которых четыре по профилю, а остальные про разработку, фриланс и
    affiliate-маркетинг. Каждый канал это 100-300 постов в месяц и плата за
    разбор каждого, поэтому брать всё подряд дороже, чем полезно.
    """
    for i, a in enumerate(sys.argv):
        if a == "--only" and i + 1 < len(sys.argv):
            return {h.strip().lstrip("@").lower()
                    for h in sys.argv[i + 1].split(",") if h.strip()}
        if a.startswith("--only="):
            return {h.strip().lstrip("@").lower()
                    for h in a.split("=", 1)[1].split(",") if h.strip()}
    return set()


def main() -> int:
    add = "--add" in sys.argv
    only = _only()
    items = asyncio.run(collect())
    if not items:
        print("новых каналов среди подписок не найдено")
        return 0
    print(f"подписок вне списка отслеживания: {len(items)} — спрашиваю модель\n")

    verdicts = judge(items)
    fits = [it for it in items if (verdicts.get(it["handle"], {})).get("fits")]
    rest = [it for it in items if it not in fits]
    if only:
        missing = only - {it["handle"] for it in fits}
        fits = [it for it in fits if it["handle"] in only]
        if missing:
            # Громко: молча взять девять из десяти названных значит оставить
            # человека в уверенности, что подписка оформлена.
            print(f"НЕ НАЙДЕНЫ среди подходящих: {', '.join(sorted(missing))}\n")

    s = get_session_factory()()
    added = 0
    try:
        for it in fits:
            v = verdicts[it["handle"]]
            niche = (v.get("niche") or "")[:100]
            print(f"  + @{it['handle']}  {it['title'][:40]}")
            print(f"      {niche or '—'} · {v.get('reason', '')[:80]}")
            if not it["known_candidate"]:
                s.add(ChannelCandidate(
                    handle=it["handle"], title=it["title"],
                    description=(v.get("reason") or "")[:500],
                    subscriber_count=it["subs"],
                    source="my_subscriptions", search_query="",
                    status="added" if add else "pending",
                ))
            if add:
                s.add(Channel(handle=it["handle"], title=it["title"],
                              niche=niche, active=True, source="my_subscriptions",
                              subscriber_count=it["subs"], verified_at=utcnow()))
                added += 1
        s.commit()
    finally:
        s.close()

    if rest:
        print(f"\nне подошли ({len(rest)}):")
        for it in rest:
            why = (verdicts.get(it["handle"], {})).get("reason") or "нет решения"
            print(f"  − @{it['handle']}: {why[:70]}")

    print(f"\nподходящих: {len(fits)}")
    if add:
        print(f"добавлено в отслеживание: {added}")
    else:
        print("это предпросмотр — повтори с --add, чтобы взять на отслеживание")
    return 0


if __name__ == "__main__":
    sys.exit(main())
