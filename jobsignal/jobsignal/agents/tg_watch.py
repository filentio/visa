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

# Боты, которым готовим черновик ответа. Не все подряд: hh_rabota_bot шлёт
# уведомления, отвечать там нечего, а ГигаРекрутер ведёт интервью.
DRAFT_BOTS = {
    h.strip().lstrip("@").lower()
    for h in (os.environ.get("TG_DRAFT_BOTS") or "giga_recruiter_bot").split(",")
    if h.strip()
}
# Файл с историями — единственный источник фактов для ответа.
STORIES = Path(os.environ.get("TG_STORIES", "config/cv_stories.md"))
# Сколько последних сообщений диалога давать для связности.
CONTEXT_MSGS = int(os.environ.get("TG_DRAFT_CONTEXT", "6"))

DRAFT_SYSTEM = (
    "Ты помогаешь кандидату отвечать рекрутёру в переписке. Пишешь ЧЕРНОВИК "
    "ответа от первого лица, который человек прочитает и отправит сам.\n\n"
    "ГЛАВНОЕ ПРАВИЛО: бери факты, числа и названия ТОЛЬКО из базы историй "
    "ниже. Ничего не добавляй от себя. Если в базе нет нужного — так и напиши "
    "одной строкой: «НЕТ ДАННЫХ: <чего не хватает>», и дальше ответь тем, что "
    "есть. Выдуманная деталь в переписке с рекрутёром хуже, чем её отсутствие: "
    "её проверят на собеседовании.\n\n"
    "Как писать: по делу, без воды и канцелярита, деловым разговорным языком. "
    "Конкретные числа из базы приводи. Если вопрос про опыт, которого нет — "
    "скажи об этом прямо и назови смежный опыт, как в разделе «честные "
    "границы». Объём — под вопрос: на короткий вопрос короткий ответ.\n\n"
    "Без приветствий в начале, если переписка уже идёт. Без подписи. "
    "Только текст сообщения.\n\n=== БАЗА ИСТОРИЙ ===\n{stories}"
)


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


def _make_draft(history: list[tuple[str, str]]) -> str | None:
    """Черновик ответа на последний вопрос. None — если нечем или не вышло."""
    if not STORIES.exists():
        log.warning("[tg_watch] нет файла историй %s — черновик не делаю", STORIES)
        return None
    try:
        stories = STORIES.read_text(encoding="utf-8")
    except Exception as exc:  # noqa: BLE001
        log.warning("[tg_watch] файл историй не читается: %s", exc)
        return None

    lines = [f"{'Я' if mine else 'Рекрутёр'}: {text}" for mine, text in history]
    user = ("Переписка (последние сообщения, снизу самое новое):\n\n"
            + "\n\n".join(lines)
            + "\n\nНапиши черновик моего ответа на последнее сообщение.")
    try:
        from jobsignal.config import get_config
        from jobsignal.llm import complete_text
        model = get_config().settings.anthropic_model
        # cache_system=True: база историй одна и та же, а она тут основной
        # объём — со второго вопроса за пять минут читается из кэша.
        return complete_text(DRAFT_SYSTEM.replace("{stories}", stories),
                             user, model=model, max_tokens=1500,
                             tag="tg_draft", cache_system=True).strip()
    except Exception as exc:  # noqa: BLE001 — черновик не критичен
        log.warning("[tg_watch] черновик не составлен: %s", exc)
        return None


def _send_draft(who: str, draft: str) -> None:
    from jobsignal.agents.notify_bot import _send

    warn = ""
    if "НЕТ ДАННЫХ" in draft:
        warn = ("\n\n⚠️ <i>В базе историй не хватает фактов — проверь "
                "отмеченное место перед отправкой.</i>")
    _send(f"✍️ <b>Черновик ответа для @{html.escape(who)}</b>\n"
          f"<i>Прочитай, поправь если надо, отправь сам.</i>\n\n"
          f"<pre>{html.escape(draft)}</pre>{warn}")


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

            # Черновик — только для интервьюирующих ботов и только на текст:
            # к сообщению с кнопками ответ печатать не нужно, там выбор.
            if uname in DRAFT_BOTS and not buttons and len(text) > 80:
                history = []
                async for m in client.iter_messages(ent, limit=CONTEXT_MSGS):
                    body = (m.message or "").strip()
                    if body:
                        history.append((bool(m.out), body))
                history.reverse()
                draft = _make_draft(history)
                if draft:
                    _send_draft(uname, draft)
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
        return asyncio.run(_run_async())
