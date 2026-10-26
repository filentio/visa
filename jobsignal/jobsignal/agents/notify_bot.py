"""
NotifyBot — Telegram-бот для уведомлений о новых вакансиях.

Два флоу в зависимости от типа контакта:

TG-контакт (@handle):
  Сообщение 1: уведомление + кнопка «Открыть чат с рекрутёром»
  Сообщение 2: готовый текст отклика + PDF резюме (пересылаешь рекрутёру)

hh.ru / форма / LinkedIn:
  Сообщение 1: уведомление + кнопка «Открыть вакансию»
              + инлайн-кнопка «✏️ Получить сопроводительное»
  При нажатии кнопки — бот присылает текст сопроводительного
  (генерация по запросу, не автоматически — экономия токенов)

Резюме прикрепляется ТОЛЬКО для TG-контактов.
Для hh.ru — не нужно (прикрепляют на сайте).
"""
from __future__ import annotations

import json
import logging
import os
import sqlite3
import time
from pathlib import Path

import requests

log = logging.getLogger("jobsignal")

DB_PATH = "data/jobsignal.db"
RESUME_DIR = Path("config/resumes")

def _resume_key(profile_name: str) -> str:
    """Ключ файла резюме по имени профиля (реестр jobsignal/profiles.py).
    Снятый или неизвестный профиль — профиль по умолчанию."""
    from jobsignal import profiles
    p = profiles.resolve(profile_name)
    return p.key if p else ""


def _token() -> str:
    return os.environ.get("NOTIFY_BOT_TOKEN", "")

def _chat() -> str:
    return os.environ.get("NOTIFY_CHAT_ID", "")

def _api(method: str) -> str:
    return f"https://api.telegram.org/bot{_token()}/{method}"


def _ensure_columns():
    conn = sqlite3.connect(DB_PATH)
    cols = [r[1] for r in conn.execute("PRAGMA table_info(vacancies)").fetchall()]
    if "notified_at" not in cols:
        conn.execute("ALTER TABLE vacancies ADD COLUMN notified_at DATETIME")
        conn.commit()
    conn.close()


def _send(text: str, markup: dict | None = None) -> int | None:
    """Send message, return message_id."""
    payload = {
        "chat_id": _chat(),
        "text": text,
        "parse_mode": "HTML",
        "disable_web_page_preview": True,
    }
    if markup:
        payload["reply_markup"] = json.dumps(markup)
    try:
        r = requests.post(_api("sendMessage"), data=payload, timeout=15)
        d = r.json()
        if d.get("ok"):
            return d["result"]["message_id"]
        log.warning("[notify] sendMessage error: %s", d.get("description"))
    except Exception as e:
        log.error("[notify] send error: %s", e)
    return None


def _send_doc(file_path: str, caption: str) -> bool:
    """Send document with caption."""
    try:
        with open(file_path, "rb") as f:
            r = requests.post(
                _api("sendDocument"),
                data={"chat_id": _chat(), "caption": caption, "parse_mode": "HTML"},
                files={"document": f},
                timeout=30,
            )
        return r.json().get("ok", False)
    except Exception as e:
        log.error("[notify] send_doc error: %s", e)
        return False


def _esc(text: str) -> str:
    return (text or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _best_profile(conn: sqlite3.Connection, vid: int) -> tuple[str, int, str]:
    """Лучший балл среди ДЕЙСТВУЮЩИХ профилей: оценки снятых профилей в базе
    остаются, но решать по ним, каким резюме откликаться, нельзя."""
    from jobsignal import profiles
    active = profiles.names()
    marks = ",".join("?" * len(active)) or "''"
    row = conn.execute(
        f"SELECT profile_key, score, reason FROM match_scores "
        f"WHERE vacancy_id=? AND profile_key IN ({marks}) "
        f"ORDER BY score DESC LIMIT 1", (vid, *active)
    ).fetchone()
    if row:
        return row[0], row[1], row[2] or ""
    return profiles.default_name(), 0, ""


def _vacancy_text_by_id(vid: int | None) -> str:
    """Полный текст вакансии по её номеру: исходный пост, иначе пересказ."""
    if not vid:
        return ""
    try:
        conn = sqlite3.connect(DB_PATH)
        row = conn.execute(
            "SELECT coalesce(p.text, v.description, '') FROM vacancies v "
            "LEFT JOIN raw_posts p ON p.id = v.raw_post_id WHERE v.id = ?",
            (vid,)).fetchone()
        conn.close()
        return (row[0] if row else "") or ""
    except Exception as e:  # noqa: BLE001 — без текста письмо всё равно пишется
        log.warning("[notify] текст вакансии #%s не прочитан: %s", vid, e)
        return ""


def _cover_text(role: str, company: str, profile_key: str, recruiter_name: str = "",
                vacancy_text: str = "", vacancy_id: int | None = None,
                channel: str = "hh", accent: str | None = None) -> str:
    """Письмо через единый сервис jobsignal/letters.py. Пустая строка — письма нет.

    Обёртка оставлена ради вызывающих (бот, кабинет, автоотклик, утилиты).
    Раньше здесь был свой промпт, а поверх него шаблон из таблицы
    cover_templates; теперь правила одни для всех путей, см. letters.py.
    Заглушку «Заинтересовала вакансия… Прилагаю резюме» при сбое больше не
    возвращаем: вызывающий должен видеть, что письма нет, а не отправить
    две строки.
    """
    from jobsignal import letters

    # Текст вакансии по номеру, если вызывающий его не передал: письмо по
    # одному названию роли безлико и подходит к любой вакансии.
    if not (vacancy_text or "").strip() and vacancy_id:
        vacancy_text = _vacancy_text_by_id(vacancy_id)
    letter = letters.generate(role, company, vacancy_text, profile=profile_key,
                              channel=channel, accent=accent,
                              recruiter_name=recruiter_name)
    return letter.text


class NotifyBot:
    def __init__(self):
        if not _token():
            raise RuntimeError("NOTIFY_BOT_TOKEN не задан")
        if not _chat():
            raise RuntimeError("NOTIFY_CHAT_ID не задан")
        _ensure_columns()

    def check_and_notify(self) -> dict:
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row

        rows = conn.execute("""
            SELECT id, role, company, salary, location, description,
                   recruiter_handle, link, contact_type
            FROM vacancies
            WHERE status = 'matched'
              AND is_primary = 1
              AND notified_at IS NULL
            ORDER BY created_at DESC
            LIMIT 20
        """).fetchall()

        sent = errors = 0
        for v in rows:
            try:
                self._notify(conn, v)
                conn.execute(
                    "UPDATE vacancies SET notified_at=datetime('now') WHERE id=?",
                    (v["id"],)
                )
                conn.commit()
                sent += 1
                time.sleep(1)
            except Exception as e:
                log.error("[notify] vacancy %s: %s", v["id"], e)
                errors += 1

        conn.close()
        log.info("[notify] отправлено: %d, ошибок: %d", sent, errors)
        return {"agent": "notify_bot", "sent": sent, "errors": errors}

    def _notify(self, conn, v):
        vid = v["id"]
        role = v["role"] or "Без названия"
        company = v["company"] or "—"
        salary = v["salary"] or ""
        location = v["location"] or ""
        desc = (v["description"] or "")[:250]
        handle = v["recruiter_handle"]
        link = v["link"] or ""
        ct = v["contact_type"] or "unknown"

        profile_key, score, reason = _best_profile(conn, vid)
        emoji = "🟢" if score >= 80 else "🟡" if score >= 70 else "⚪"

        # ── notification text ─────────────────────────────────────────────────
        text = (
            f"{emoji} <b>{_esc(role)}</b>\n"
            f"🏢 {_esc(company)}\n"
        )
        if salary:
            text += f"💰 {_esc(salary)}\n"
        if location:
            text += f"📍 {_esc(location)}\n"
        text += f"📊 Балл: <b>{score}</b> ({_esc(profile_key)})\n"
        if reason:
            text += f"<i>{_esc(reason[:120])}</i>\n"
        if desc:
            text += f"\n{_esc(desc)}"

        if ct == "tg" and handle:
            # ── TG flow: open chat button ──────────────────────────────────
            username = handle.lstrip("@")
            markup = {"inline_keyboard": [
                [{"text": "💬 Открыть чат с рекрутёром",
                  "url": f"https://t.me/{username}"}],
                [{"text": "✅ Откликнулся",
                  "callback_data": f"applied:{vid}"}],
            ]}
            _send(text, markup)

            # second message: forward-ready with resume
            cover = _cover_text(role, company, profile_key, vacancy_id=vid,
                                channel="telegram")
            caption = (f"👇 <i>Перешли это сообщение рекрутёру</i>\n\n{_esc(cover)}"
                       if cover else "⚠️ Письмо не получилось (модель не ответила "
                       "или нет резюме профиля) — резюме прикладываю без текста")
            from jobsignal.agents.resume_parser import pdf_path_for
            resume_path = pdf_path_for(_resume_key(profile_key))

            if resume_path is not None:
                _send_doc(str(resume_path), caption)
            else:
                _send(caption)

        else:
            # ── hh / form / linkedin flow: open link + on-demand cover ────
            link_label = {
                "hh": "🔗 Открыть на hh.ru",
                "form": "🔗 Открыть форму",
                "linkedin": "🔗 Открыть LinkedIn",
            }.get(ct, "🔗 Открыть вакансию")

            buttons = []
            if link:
                buttons.append({"text": link_label, "url": link})

            # callback button to generate cover on demand
            # we encode vacancy_id in callback_data
            buttons.append({
                "text": "✏️ Получить сопроводительное",
                "callback_data": f"cover:{vid}:{profile_key}"
            })

            applied_btn = [{"text": "✅ Откликнулся",
                            "callback_data": f"applied:{vid}"}]
            markup = {"inline_keyboard": [buttons, applied_btn]} if buttons else None
            _send(text, markup)

    def handle_callback(self, callback_query: dict) -> None:
        """Handle inline button press — generate cover letter on demand."""
        cq_id = callback_query.get("id")
        data = callback_query.get("data", "")
        chat_id = callback_query.get("message", {}).get("chat", {}).get("id")

        # answer callback immediately (removes loading spinner in TG)
        try:
            requests.post(_api("answerCallbackQuery"),
                         data={"callback_query_id": cq_id}, timeout=5)
        except Exception:
            pass

        # handle "applied" button
        if data.startswith("applied:"):
            vid = int(data.split(":")[1])
            conn = sqlite3.connect(DB_PATH)
            conn.execute("UPDATE vacancies SET status='applied' WHERE id=?", (vid,))
            row = conn.execute(
                "SELECT role, company, draft_text FROM vacancies WHERE id=?",
                (vid,)
            ).fetchone()
            # create application record
            from datetime import datetime, timezone
            # Формат как у SQLAlchemy и datetime('now'): «ГГГГ-ММ-ДД ЧЧ:ММ:СС».
            # isoformat() давал «…T01:00:00+00:00», а квота сравнивает время
            # строками: «T» больше пробела, и ручная отметка считалась «в
            # пределах часа» до конца суток, съедая квоту автоотклика.
            now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S.%f")
            # Кнопку жмёт человек и отклик отправляет сам, поэтому канал —
            # manual, а не telegram: иначе ручные отметки неотличимы от писем,
            # которые действительно отправила система.
            # Раньше здесь стоял INSERT OR IGNORE только с (vacancy_id,
            # sent_at): message_text и channel в схеме NOT NULL, а слово
            # IGNORE гасит именно это нарушение — строка молча не
            # создавалась, отклик не попадал ни в квоту, ни в аналитику.
            already = conn.execute(
                "SELECT 1 FROM applications WHERE vacancy_id=? AND sent_at IS NOT NULL",
                (vid,)
            ).fetchone()
            if not already:
                conn.execute(
                    "INSERT INTO applications "
                    "(vacancy_id, message_text, channel, is_draft, sent_at, created_at) "
                    "VALUES (?,?,'manual',0,?,?)",
                    (vid, (row[2] or "").strip() if row else "", now, now)
                )
            conn.commit()
            conn.close()
            role = row[0] if row else "вакансия"
            company = row[1] if row else ""
            try:
                requests.post(_api("sendMessage"), data={
                    "chat_id": chat_id,
                    "text": f"✅ Отмечено: откликнулся на <b>{_esc(role)}</b>{(' · ' + _esc(company)) if company else ''}",
                    "parse_mode": "HTML",
                }, timeout=10)
            except Exception:
                pass
            return

        # Кнопка «Пропустить» у телеграм-откликов: человек посмотрел и решил
        # не писать. Без неё вакансия возвращалась бы в очередь каждый прогон.
        if data.startswith("tgskip:"):
            vid = int(data.split(":")[1])
            conn = sqlite3.connect(DB_PATH)
            conn.execute("UPDATE vacancies SET status='skipped' WHERE id=?", (vid,))
            row = conn.execute("SELECT role, company FROM vacancies WHERE id=?",
                               (vid,)).fetchone()
            conn.commit()
            conn.close()
            role = row[0] if row else "вакансия"
            try:
                requests.post(_api("sendMessage"), data={
                    "chat_id": chat_id,
                    "text": f"⏭ Пропущено: <b>{_esc(role)}</b>",
                    "parse_mode": "HTML",
                }, timeout=10)
            except Exception:
                pass
            return

        if not data.startswith("cover:"):
            return

        parts = data.split(":", 2)
        if len(parts) < 3:
            return
        vid = int(parts[1])
        profile_key = parts[2]

        conn = sqlite3.connect(DB_PATH)
        row = conn.execute(
            "SELECT role, company, contact_type FROM vacancies WHERE id=?", (vid,)
        ).fetchone()
        conn.close()

        if not row:
            return

        role, company, ct = row
        # Канал по контакту вакансии: правила письма для hh, анкеты и почты разные.
        channel = {"hh": "hh", "form": "form", "email": "email",
                   "tg": "telegram"}.get(ct or "", "hh")
        cover = _cover_text(role or "", company or "", profile_key, vacancy_id=vid,
                            channel=channel) or "Письмо не получилось — попробуй ещё раз."

        msg = (
            f"📋 <b>Сопроводительное письмо</b>\n\n"
            f"{_esc(cover)}\n\n"
            f"<i>Скопируй и вставь в форму отклика</i>"
        )
        try:
            requests.post(_api("sendMessage"), data={
                "chat_id": chat_id,
                "text": msg,
                "parse_mode": "HTML",
            }, timeout=15)
        except Exception as e:
            log.error("[notify] callback send error: %s", e)

    def run_webhook_polling(self, interval: int = 5) -> None:
        """
        Long-poll for callback queries (button presses).
        Run this in background to handle '✏️ Получить сопроводительное' button.
        """
        offset = 0
        log.info("[notify] polling for callbacks...")
        while True:
            try:
                r = requests.get(
                    _api("getUpdates"),
                    # allowed_updates Telegram ждёт строкой JSON: список из
                    # requests уходил повторяющимися параметрами.
                    params={"offset": offset, "timeout": 30,
                            "allowed_updates": json.dumps(["callback_query"])},
                    timeout=35,
                )
                data = r.json()
                if not data.get("ok"):
                    # 409 — второй экземпляр бота или установлен webhook, 401 —
                    # отозван токен. Раньше это молча давало пустой список, и
                    # цикл без паузы долбил Telegram, а кнопки не работали.
                    log.error("[notify] getUpdates: %s %s — пауза %d с",
                              data.get("error_code"), data.get("description"),
                              max(interval, 30))
                    time.sleep(max(interval, 30))
                    continue
                updates = data.get("result", [])
                for upd in updates:
                    offset = upd["update_id"] + 1
                    if "callback_query" in upd:
                        self.handle_callback(upd["callback_query"])
            except Exception as e:
                log.error("[notify] polling error: %s", e)
                time.sleep(interval)
