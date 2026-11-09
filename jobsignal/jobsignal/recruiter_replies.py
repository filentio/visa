"""Ответы рекрутёрам в Telegram с согласованием: черновик → кнопка → отправка.

Зачем. 09.11 ГигаРекрутер Сбера вёл интервью из шести вопросов, и каждый
ответ человек пересылал в чат, получал текст, копировал и отправлял. Теперь
сторож входящих (agents/tg_watch.py) готовит черновик сам и присылает его в
бот с кнопками:

  ✅ Отправить  — ответ уходит с аккаунта кандидата (та же сессия Telethon);
  ✏️ Изменить   — человек отвечает на черновик (Reply) своим текстом, бот
                  показывает новый вариант с теми же кнопками;
  ⏭ Пропустить  — ничего не отправляем.

Факты для ответа — только из резюме профилей (без контактов) и, если есть,
config/cv_stories.md. Чего в них нет (зарплата, причины ухода, личные
примеры) — модель помечает «НЕТ ДАННЫХ: …», и такой черновик приходит с
предупреждением: его нужно дописать через «Изменить».

Защита от устаревшего ответа: перед отправкой проверяем, что после вопроса
собеседник ничего не написал. Если написал — не отправляем: на новое
сообщение придёт новый черновик.

Черновики — data/tg_drafts.json: {id: {who, peer_msg_id, text, status,
bot_msg_id, at}}.
"""
from __future__ import annotations

import asyncio
import json
import logging
import os
import secrets
import time
from pathlib import Path

log = logging.getLogger("jobsignal")

DRAFTS = Path(os.environ.get("TG_DRAFTS_FILE", "data/tg_drafts.json"))
STORIES = Path(os.environ.get("TG_STORIES", "config/cv_stories.md"))

SYSTEM = (
    "Ты помогаешь кандидату отвечать рекрутёру (или боту-рекрутёру) в переписке. "
    "Пишешь ответ от первого лица, который кандидат согласует и отправит.\n\n"
    "ГЛАВНОЕ ПРАВИЛО: факты, числа, названия компаний и проектов — ТОЛЬКО из "
    "резюме и историй ниже. Ничего не добавляй. Если для ответа нужно то, чего "
    "там нет (зарплатные ожидания, причины ухода, конкретные истории про людей, "
    "график, личные обстоятельства), ответь тем, что есть, и в конце отдельной "
    "строкой напиши «НЕТ ДАННЫХ: <чего не хватает>». Выдуманная деталь хуже "
    "отсутствующей: её проверят на собеседовании.\n\n"
    "Как писать: по делу, живым деловым языком, без канцелярита и воды. Объём — "
    "под вопрос: на короткий вопрос — короткий ответ, на просьбу рассказать о "
    "проекте — 4-8 предложений с цифрами из резюме. Выбирай опыт, который ближе "
    "к вакансии из переписки. Без приветствия, если переписка уже идёт, без "
    "подписи, без markdown. Только текст сообщения.\n\n=== РЕЗЮМЕ И ИСТОРИИ ===\n{facts}"
)


def _load() -> dict:
    try:
        return json.loads(DRAFTS.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def _save(data: dict) -> None:
    # Старше двух недель — вычищаем.
    cutoff = time.time() - 14 * 86400
    data = {k: v for k, v in data.items() if v.get("at", 0) >= cutoff}
    DRAFTS.parent.mkdir(parents=True, exist_ok=True)
    tmp = DRAFTS.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(DRAFTS)


def facts() -> str:
    """Резюме всех профилей без контактов + истории, если есть."""
    from jobsignal import profiles
    from jobsignal.agents.matcher import strip_contacts
    parts = []
    for p in profiles.all_profiles():
        try:
            text = strip_contacts(p.cv_text())
        except Exception as exc:  # noqa: BLE001
            log.warning("[replies] резюме %s не прочитано: %s", p.key, exc)
            continue
        if text.strip():
            parts.append(f"--- Резюме «{p.name}» ---\n{text[:9000]}")
    if STORIES.exists():
        try:
            parts.append("--- Истории ---\n" + STORIES.read_text(encoding="utf-8")[:12000])
        except OSError:
            pass
    return "\n\n".join(parts)


def make_draft(history: list[tuple[bool, str]]) -> str | None:
    """Черновик ответа на последнее сообщение. history: [(моё?, текст)], снизу новое."""
    base = facts()
    if not base:
        log.warning("[replies] нет ни резюме, ни историй — черновик не делаю")
        return None
    lines = [f"{'Я' if mine else 'Собеседник'}: {text}" for mine, text in history]
    user = ("Переписка (снизу самое новое):\n\n" + "\n\n".join(lines)
            + "\n\nНапиши мой ответ на последнее сообщение собеседника.")
    try:
        from jobsignal.config import get_config
        from jobsignal.llm import complete_text
        return complete_text(SYSTEM.replace("{facts}", base), user,
                             model=get_config().settings.anthropic_model, max_tokens=1200,
                             tag="tg_reply", cache_system=True, task="letter").strip()
    except Exception as exc:  # noqa: BLE001
        log.warning("[replies] черновик не составлен: %s", exc)
        return None


def _markup(did: str) -> dict:
    return {"inline_keyboard": [[
        {"text": "✅ Отправить", "callback_data": f"rsend:{did}"},
        {"text": "✏️ Изменить", "callback_data": f"redit:{did}"},
        {"text": "⏭ Пропустить", "callback_data": f"rskip:{did}"},
    ]]}


def _show(did: str, d: dict, head: str = "") -> int | None:
    import html
    from jobsignal.agents.notify_bot import _send
    warn = ("\n\n⚠️ <i>В резюме не хватило фактов (строка «НЕТ ДАННЫХ») — "
            "допишите через «✏️ Изменить», иначе так и уйдёт.</i>"
            if "НЕТ ДАННЫХ" in d["text"] else "")
    text = (f"{head or '✍️ <b>Ответ для @' + html.escape(d['who']) + '</b>'}\n"
            f"<i>«Отправить» — уйдёт с вашего аккаунта. «Изменить» — ответьте на это "
            f"сообщение (Reply) своим текстом.</i>\n\n"
            f"{html.escape(d['text'])}{warn}")
    return _send(text[:4000], _markup(did))


def offer(who: str, peer_msg_id: int, draft: str) -> None:
    data = _load()
    did = secrets.token_hex(4)
    d = {"who": who, "peer_msg_id": peer_msg_id, "text": draft, "status": "offered",
         "at": time.time()}
    d["bot_msg_id"] = _show(did, d)
    data[did] = d
    _save(data)


def by_bot_message(bot_msg_id: int) -> str | None:
    for did, d in _load().items():
        if d.get("bot_msg_id") == bot_msg_id:
            return did
    return None


def edit(did: str, new_text: str) -> str:
    data = _load()
    d = data.get(did)
    if not d or d.get("status") != "offered":
        return "этот черновик уже отправлен или закрыт"
    d["text"] = new_text.strip()[:3500]
    d["bot_msg_id"] = _show(did, d, head=f"✏️ <b>Новый вариант для @{d['who']}</b>")
    data[did] = d
    _save(data)
    return ""


def skip(did: str) -> str:
    data = _load()
    d = data.get(did)
    if not d:
        return "черновик не найден"
    d["status"] = "skipped"
    _save(data)
    return f"пропущено, @{d['who']} ничего не отправлено"


async def _send_async(who: str, peer_msg_id: int, text: str) -> None:
    from telethon import TelegramClient
    from jobsignal.tg_send import TGSendError

    api_id = int(os.environ.get("TG_API_ID") or 0)
    api_hash = (os.environ.get("TG_API_HASH") or "").strip()
    session = (os.environ.get("TG_INBOX_SESSION") or "data/tg_inbox").strip()
    client = TelegramClient(session, api_id, api_hash)
    await client.connect()
    try:
        if not await client.is_user_authorized():
            raise TGSendError("сессия Telegram на сервере не авторизована")
        entity = await client.get_entity("@" + who)
        last = await client.get_messages(entity, limit=1)
        if last and not last[0].out and last[0].id != peer_msg_id:
            raise TGSendError("собеседник уже написал новое сообщение — этот ответ "
                              "устарел, на новое придёт свой черновик")
        await client.send_message(entity, text)
    finally:
        await client.disconnect()


def send(did: str) -> str:
    """Отправить согласованный ответ. Возвращает текст для человека."""
    from jobsignal.tg_lock import tg_session_lock
    from jobsignal.tg_send import TGSendError
    data = _load()
    d = data.get(did)
    if not d:
        return "⚠️ черновик не найден"
    if d.get("status") == "sent":
        return "ℹ️ уже отправлено"
    if d.get("status") == "skipped":
        return "ℹ️ этот черновик пропущен"
    if "НЕТ ДАННЫХ" in d["text"]:
        return ("⚠️ в ответе осталась строка «НЕТ ДАННЫХ» — допишите через "
                "«✏️ Изменить», потом «Отправить»")
    try:
        with tg_session_lock(wait_s=90) as got:
            if not got:
                return "⚠️ сессия Telegram занята — нажмите ещё раз через минуту"
            asyncio.run(_send_async(d["who"], d["peer_msg_id"], d["text"]))
    except TGSendError as exc:
        return f"⚠️ не отправлено: {exc}"
    except Exception as exc:  # noqa: BLE001
        log.error("[replies] отправка %s: %s", did, exc)
        return f"⚠️ не отправлено: {type(exc).__name__}: {exc}"
    d["status"] = "sent"
    d["sent_at"] = time.time()
    _save(data)
    log.info("[replies] ответ отправлен @%s", d["who"])
    return f"✅ отправлено @{d['who']}"
