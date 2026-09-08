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

PROFILE_KEY_MAP = {
    "Senior AI PM": "ai_pm",
    "CPO / Head of Product": "cpo",
    "Senior PM/PO": "pm",
}


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
    row = conn.execute(
        "SELECT profile_key, score, reason FROM match_scores "
        "WHERE vacancy_id=? ORDER BY score DESC LIMIT 1", (vid,)
    ).fetchone()
    if row:
        return row[0], row[1], row[2] or ""
    return "Senior PM/PO", 0, ""


def _get_active_template() -> dict | None:
    """Get active cover letter template from DB."""
    try:
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        row = conn.execute(
            "SELECT * FROM cover_templates WHERE is_active=1 LIMIT 1"
        ).fetchone()
        conn.close()
        return dict(row) if row else None
    except Exception:
        return None


def _cover_text(role: str, company: str, profile_key: str, recruiter_name: str = "") -> str:
    """Generate cover letter using active template + real resume text. No hallucinations."""
    from jobsignal.agents.resume_parser import get_resume_text
    from jobsignal.llm import complete_text

    resume_key = PROFILE_KEY_MAP.get(profile_key, "pm")
    resume_text = get_resume_text(resume_key)

    if not resume_text:
        from jobsignal.agents.resume_parser import parse_and_save
        result = parse_and_save(resume_key)
        if result.get("ok"):
            resume_text = get_resume_text(resume_key)

    company_part = f" в {company}" if company and company not in ("—", "None", "") else ""
    name_part = f", {recruiter_name}" if recruiter_name else ""

    if not resume_text:
        return f"Добрый день{name_part}\n\nЗаинтересовала вакансия {role}{company_part}. Прилагаю резюме — готов обсудить."

    # get active template
    tpl = _get_active_template()
    system_prompt = (
        tpl["system_prompt"] if tpl and tpl.get("system_prompt")
        else (
            "Напиши короткое сопроводительное письмо (5-7 предложений). "
            "Используй ТОЛЬКО факты из резюме — никаких выдуманных цифр и компаний. "
            "Стиль: деловой, живой, как человек пишет человеку. "
            "Без шаблонных фраз типа \"уверен что мой опыт\", \"стремлюсь к развитию\". "
            "Структура: приветствие → вакансия → кто я + компании → 1-2 результата → призыв обсудить."
        )
    )

    profile_labels = {
        "Senior AI PM": "AI Product Owner",
        "CPO / Head of Product": "CPO / Head of Product",
        "Senior PM/PO": "Senior Product Manager",
    }
    companies_by_profile = {
        "ai_pm": "ФИНАМ, ОТП Банк",
        "cpo": "ФИНАМ, Мосбиржа",
        "pm": "ФИНАМ, Мосбиржа, ОТП Банк",
    }

    user = (
        f"Вакансия: {role}{company_part}\n"
        f"Имя рекрутёра: {recruiter_name or 'неизвестно'}\n"
        f"Профиль кандидата: {profile_labels.get(profile_key, profile_key)}\n"
        f"Компании из опыта: {companies_by_profile.get(resume_key, 'ФИНАМ, Мосбиржа')}\n\n"
        f"Резюме кандидата:\n{resume_text[:3500]}\n\n"
        f"Напиши сопроводительное письмо. "
        f"Начни с \'Добрый день{name_part}\'. "
        f"Только факты из резюме, никаких выдумок."
    )

    try:
        return complete_text(system_prompt, user, "", 350)
    except Exception as e:
        log.warning("[notify] cover LLM error: %s", e)
        return f"Добрый день{name_part}\n\nЗаинтересовала вакансия {role}{company_part}. Прилагаю резюме — готов обсудить."


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
            cover = _cover_text(role, company, profile_key)
            caption = f"👇 <i>Перешли это сообщение рекрутёру</i>\n\n{_esc(cover)}"
            resume_key = PROFILE_KEY_MAP.get(profile_key, "pm")
            resume_path = RESUME_DIR / f"{resume_key}.pdf"

            if resume_path.exists():
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
            # create application record
            from datetime import datetime, timezone
            now = datetime.now(timezone.utc).isoformat()
            conn.execute(
                "INSERT OR IGNORE INTO applications (vacancy_id, sent_at) VALUES (?,?)",
                (vid, now)
            )
            conn.commit()
            # get vacancy name for confirmation
            row = conn.execute("SELECT role, company FROM vacancies WHERE id=?", (vid,)).fetchone()
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

        if not data.startswith("cover:"):
            return

        parts = data.split(":", 2)
        if len(parts) < 3:
            return
        vid = int(parts[1])
        profile_key = parts[2]

        conn = sqlite3.connect(DB_PATH)
        row = conn.execute(
            "SELECT role, company FROM vacancies WHERE id=?", (vid,)
        ).fetchone()
        conn.close()

        if not row:
            return

        role, company = row
        cover = _cover_text(role or "", company or "", profile_key)

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
                    params={"offset": offset, "timeout": 30,
                            "allowed_updates": ["callback_query"]},
                    timeout=35,
                )
                updates = r.json().get("result", [])
                for upd in updates:
                    offset = upd["update_id"] + 1
                    if "callback_query" in upd:
                        self.handle_callback(upd["callback_query"])
            except Exception as e:
                log.error("[notify] polling error: %s", e)
                time.sleep(interval)
