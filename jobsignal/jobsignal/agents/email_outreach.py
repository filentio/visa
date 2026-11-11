"""Отклики по почте: письмо под вакансию + PDF резюме с ящика кандидата.

Зачем. У ~16 свежих вакансий в день единственный контакт — почта рекрутёра
(contact_type = email). До 09.11 по ним приходило только уведомление.

Режим — EMAIL_OUTREACH_MODE:
  off      (по умолчанию) ничего не делаем, такие вакансии идут обычным
           уведомлением;
  confirm  письмо приходит в телеграм с кнопками «📧 Отправить» / «⏭
           Пропустить», уходит по кнопке;
  auto     письмо уходит само в пределах лимитов, в телеграм — отчёт.

Очередь: matched/drafted, contact_type = email, есть recruiter_email, лучший
балл действующего профиля ≥ EMAIL_THRESHOLD (75), свежая по дате публикации,
отклика ещё не было, на этот адрес не писали EMAIL_SAME_ADDRESS_DAYS (60) дней.
Лимиты: EMAIL_PER_HOUR (3), EMAIL_PER_DAY (10) — считаются по отправленным
письмам (applications.channel = 'email').

Письмо — letters.generate(channel="email") с проверкой фактов: письмо с
фактами не из резюме не уходит. Подпись — EMAIL_FROM_NAME и EMAIL_SIGNATURE
(строки через «|»). Вложение — PDF резюме профиля.

Ящик: SMTP_USER и SMTP_PASSWORD (пароль приложения); SMTP_HOST/SMTP_PORT —
если адрес не на Яндексе, Gmail или Mail.ru.
"""
from __future__ import annotations

import logging
import mimetypes
import os
import smtplib
import ssl
from datetime import datetime, timedelta, timezone
from email.message import EmailMessage
from email.utils import formataddr, make_msgid
from pathlib import Path

from sqlalchemy import func, select

from jobsignal.db import Application, MatchScore, Vacancy, VacancyStatus, get_session_factory

log = logging.getLogger("jobsignal")

PRESETS = {
    ("yandex.ru", "ya.ru", "yandex.com", "narod.ru"): ("smtp.yandex.ru", 465),
    ("gmail.com", "googlemail.com"): ("smtp.gmail.com", 465),
    ("mail.ru", "bk.ru", "list.ru", "inbox.ru", "internet.ru"): ("smtp.mail.ru", 465),
}


def mode() -> str:
    m = (os.environ.get("EMAIL_OUTREACH_MODE") or "off").strip().lower()
    return m if m in ("off", "confirm", "auto") else "off"


def _int(name: str, default: int) -> int:
    try:
        return int(os.environ.get(name, default))
    except ValueError:
        return default


def smtp_conf() -> dict:
    user = (os.environ.get("SMTP_USER") or "").strip()
    host = (os.environ.get("SMTP_HOST") or "").strip()
    port = _int("SMTP_PORT", 0)
    if not host and "@" in user:
        domain = user.rsplit("@", 1)[1].lower()
        for domains, (h, p) in PRESETS.items():
            if domain in domains:
                host, port = h, port or p
    return {"user": user, "password": os.environ.get("SMTP_PASSWORD") or "",
            "host": host, "port": port or 465,
            "name": (os.environ.get("EMAIL_FROM_NAME") or "").strip()}


def _connect(conf: dict):
    ctx = ssl.create_default_context()
    if conf["port"] == 465:
        s = smtplib.SMTP_SSL(conf["host"], conf["port"], context=ctx, timeout=30)
    else:
        s = smtplib.SMTP(conf["host"], conf["port"], timeout=30)
        s.starttls(context=ctx)
    s.login(conf["user"], conf["password"])
    return s


def check(send_test: bool = False) -> str:
    """Проверка ящика: вход по SMTP; с send_test — письмо самому себе."""
    conf = smtp_conf()
    if not conf["user"] or not conf["password"]:
        return "не заданы SMTP_USER и SMTP_PASSWORD"
    if not conf["host"]:
        return "сервер почты не определился по адресу — задай SMTP_HOST и SMTP_PORT"
    try:
        s = _connect(conf)
    except smtplib.SMTPAuthenticationError as exc:
        return (f"ящик не пустил ({exc.smtp_code}): нужен пароль ПРИЛОЖЕНИЯ, а не "
                "обычный пароль, и включённый доступ по SMTP/IMAP в настройках почты")
    except Exception as exc:  # noqa: BLE001
        return f"не подключился к {conf['host']}:{conf['port']}: {type(exc).__name__}: {exc}"
    try:
        if send_test:
            msg = EmailMessage()
            msg["From"] = formataddr((conf["name"] or conf["user"], conf["user"]))
            msg["To"] = conf["user"]
            msg["Subject"] = "jobsignal: проверка отправки"
            msg.set_content("Это проверочное письмо: отправка откликов по почте настроена.")
            s.send_message(msg)
            return f"ok: вход в {conf['host']} есть, проверочное письмо отправлено на {conf['user']}"
        return f"ok: вход в {conf['host']} есть"
    finally:
        try:
            s.quit()
        except Exception:  # noqa: BLE001
            pass


def _signature() -> str:
    name = (os.environ.get("EMAIL_FROM_NAME") or "").strip()
    extra = [x.strip() for x in (os.environ.get("EMAIL_SIGNATURE") or "").split("|") if x.strip()]
    lines = ["С уважением,"] + ([name] if name else []) + extra
    return "\n".join(lines) if (name or extra) else ""


def _subject(role: str) -> str:
    name = (os.environ.get("EMAIL_FROM_NAME") or "").strip()
    base = f"Отклик на вакансию «{(role or 'вакансия').strip()}»"
    return f"{base} — {name}" if name else base


def send(to: str, subject: str, body: str, attachment: Path | None) -> str:
    """Отправить письмо. Возвращает Message-ID."""
    conf = smtp_conf()
    msg = EmailMessage()
    msg["From"] = formataddr((conf["name"] or conf["user"], conf["user"]))
    msg["To"] = to
    msg["Reply-To"] = conf["user"]
    msg["Subject"] = subject
    mid = make_msgid(domain=conf["user"].rsplit("@", 1)[-1] or None)
    msg["Message-ID"] = mid
    msg.set_content(body)
    if attachment is not None and attachment.exists():
        ctype = mimetypes.guess_type(attachment.name)[0] or "application/pdf"
        main, sub = ctype.split("/", 1)
        # «po.pdf» в письме работодателю выглядит небрежно: имя файла — по
        # имени кандидата («Резюме Павел Прищутов.pdf»).
        name = (os.environ.get("EMAIL_FROM_NAME") or "").strip()
        fname = (os.environ.get("EMAIL_RESUME_FILENAME")
                 or (f"Резюме {name}{attachment.suffix}" if name else attachment.name))
        msg.add_attachment(attachment.read_bytes(), maintype=main, subtype=sub,
                           filename=fname)
    s = _connect(conf)
    try:
        s.send_message(msg)
    finally:
        try:
            s.quit()
        except Exception:  # noqa: BLE001
            pass
    return mid


class EmailOutreach:
    name = "email_outreach"

    def __init__(self):
        self._sf = get_session_factory()

    # --- очередь и лимиты --------------------------------------------------

    def _sent_since(self, s, hours: float) -> int:
        since = datetime.now(timezone.utc).replace(tzinfo=None) - timedelta(hours=hours)
        return s.execute(select(func.count()).select_from(Application).where(
            Application.channel == "email", Application.sent_at.isnot(None),
            Application.sent_at >= since)).scalar() or 0

    def allowed_now(self, s) -> int:
        return max(0, min(_int("EMAIL_PER_HOUR", 3) - self._sent_since(s, 1),
                          _int("EMAIL_PER_DAY", 10) - self._sent_since(s, 24)))

    def _recently_emailed(self, s, address: str) -> bool:
        since = (datetime.now(timezone.utc).replace(tzinfo=None)
                 - timedelta(days=_int("EMAIL_SAME_ADDRESS_DAYS", 60)))
        return bool(s.execute(
            select(func.count()).select_from(Application).join(Vacancy)
            .where(Application.channel == "email", Application.sent_at.isnot(None),
                   Application.sent_at >= since,
                   func.lower(Vacancy.recruiter_email) == address.lower())).scalar())

    def queue(self, s) -> list[tuple[Vacancy, int, str]]:
        from jobsignal import freshness, profiles
        active = set(profiles.names())
        threshold = _int("EMAIL_THRESHOLD", 75)
        rows = s.execute(select(Vacancy).where(
            Vacancy.is_primary.is_(True), Vacancy.contact_type == "email",
            Vacancy.recruiter_email.isnot(None),
            Vacancy.status.in_([VacancyStatus.matched, VacancyStatus.drafted]),
            ~Vacancy.applications.any(Application.sent_at.isnot(None)),
        ).order_by(Vacancy.id.desc())).scalars().all()
        out = []
        for v in rows:
            if not freshness.is_fresh(v):
                continue
            best = max(((m.score or 0, m.profile_key) for m in v.match_scores
                        if m.profile_key in active), default=(0, ""))
            if best[0] >= threshold:
                out.append((v, best[0], best[1]))
        out.sort(key=lambda x: (-x[1], -x[0].id))
        return out

    # --- письмо ------------------------------------------------------------

    def _letter(self, s, v: Vacancy, profile: str):
        from jobsignal import letters
        draft = letters.reusable(v, "email", profile)
        if draft:
            return draft, []
        vac_text = ((v.raw_post.text if v.raw_post is not None else "")
                    or v.description or "")
        letter = None
        for _ in range(2):
            letter = letters.generate(v.role or "", v.company or "", vac_text,
                                      profile=profile, channel="email")
            if letter.ok and letter.facts_ok:
                letters.remember(v, letter)
                s.commit()
                return letter.text, []
        return "", (letter.issues if letter else ["письмо не получилось"])

    def _record(self, s, v: Vacancy, body: str) -> None:
        now = datetime.now(timezone.utc).replace(tzinfo=None)
        for a in v.applications:
            if a.is_draft and a.channel == "email":
                s.delete(a)
        s.add(Application(vacancy_id=v.id, message_text=body, channel="email",
                          is_draft=False, sent_at=now, created_at=now))
        v.status = VacancyStatus.applied
        s.commit()

    def send_one(self, vacancy_id: int) -> str:
        """Отправить письмо по вакансии (кнопка в телеграме или авторежим)."""
        from jobsignal import profiles
        s = self._sf()
        try:
            v = s.get(Vacancy, vacancy_id)
            if v is None or not v.recruiter_email:
                return "вакансия не найдена или у неё нет почты"
            if any(a.sent_at for a in v.applications):
                return "по этой вакансии отклик уже был"
            if self._recently_emailed(s, v.recruiter_email):
                return f"на {v.recruiter_email} уже писали недавно — повтор не отправляю"
            if self.allowed_now(s) <= 0:
                return "лимит писем на час или сутки исчерпан — попробуй позже"
            active = set(profiles.names())
            best = max(((m.score or 0, m.profile_key) for m in v.match_scores
                        if m.profile_key in active), default=(0, ""))
            prof = profiles.resolve(best[1])
            text, issues = self._letter(s, v, prof.name if prof else best[1])
            if not text:
                return "письмо не прошло проверку: " + "; ".join(issues)[:300]
            sig = _signature()
            body = text + (f"\n\n{sig}" if sig else "")
            pdf = prof.resume_pdf() if prof else None
            send(v.recruiter_email, _subject(v.role or ""), body, pdf)
            self._record(s, v, body)
            log.info("[email] #%s: письмо отправлено на %s", v.id, v.recruiter_email)
            return f"отправлено на {v.recruiter_email}" + ("" if pdf else " (без PDF резюме)")
        except smtplib.SMTPException as exc:
            log.error("[email] #%s: почта не отправила: %s", vacancy_id, exc)
            return f"почта не отправила: {exc}"
        finally:
            s.close()

    def offer(self, vacancy_id: int, text: str | None = None) -> str:
        """Письмо по одной вакансии — в бот с кнопкой «📧 Отправить» (кнопка
        «В бот» в кабинете). text — готовый текст из кабинета; пусто — пишем."""
        from jobsignal import letters, profiles
        from .notify_bot import _esc, _send
        s = self._sf()
        try:
            v = s.get(Vacancy, vacancy_id)
            if v is None or not v.recruiter_email:
                return "у вакансии нет почты рекрутёра"
            if any(a.sent_at for a in v.applications):
                return "по этой вакансии отклик уже был"
            active = set(profiles.names())
            best = max(((m.score or 0, m.profile_key) for m in v.match_scores
                        if m.profile_key in active), default=(0, ""))
            prof = profiles.resolve(best[1])
            pname = prof.name if prof else best[1]
            if text and text.strip():
                letters.remember(v, letters.Letter(text.strip(), "email", pname, "auto"))
                s.commit()
                body_text = text.strip()
            else:
                body_text, issues = self._letter(s, v, pname)
                if not body_text:
                    return "письмо не прошло проверку: " + "; ".join(issues)[:300]
            sig = _signature()
            body = body_text + (f"\n\n{sig}" if sig else "")
            markup = {"inline_keyboard": [[
                {"text": "📧 Отправить", "callback_data": f"emailsend:{v.id}"},
                {"text": "⏭ Пропустить", "callback_data": f"tgskip:{v.id}"},
            ]]}
            msg = (f"✉️ <b>Отклик по почте</b> · {best[0]}%\n"
                   f"{_esc(v.role or '—')} — {_esc(v.company or '—')}\n"
                   f"Кому: <code>{_esc(v.recruiter_email)}</code>\n"
                   f"Тема: {_esc(_subject(v.role or ''))}\n\n{_esc(body)}")
            if _send(msg[:4000], markup) is None:
                return "бот не принял сообщение"
            return "ok"
        finally:
            s.close()

    # --- прогон ------------------------------------------------------------

    def run(self) -> dict:
        m = mode()
        if m == "off":
            return {"agent": self.name, "mode": m}
        conf = smtp_conf()
        if not conf["user"] or not conf["password"] or not conf["host"]:
            log.warning("[email] режим %s, но ящик не настроен (SMTP_USER/SMTP_PASSWORD)", m)
            return {"agent": self.name, "mode": m, "error": "ящик не настроен"}
        from .notify_bot import _esc, _send
        s = self._sf()
        sent = offered = 0
        q: list = []
        try:
            q = self.queue(s)
            if m == "auto":
                for v, score, _ in q:
                    if self.allowed_now(s) <= 0:
                        break
                    if self._recently_emailed(s, v.recruiter_email):
                        continue
                    res = self.send_one(v.id)
                    ok = res.startswith("отправлено")
                    sent += ok
                    _send(("📧 <b>Письмо отправлено</b>\n" if ok else "⚠️ <b>Письмо не ушло</b>\n")
                          + f"{_esc(v.role or '—')} — {_esc(v.company or '—')} · {score}%\n"
                          + _esc(res))
                    if not ok and "лимит" in res:
                        break
            else:   # confirm: предложить в телеграме тем, кому ещё не предлагали
                per_run = _int("EMAIL_OFFERS_PER_RUN", 5)
                offered_to: set[str] = set()
                for v, score, prof_name in q:
                    if offered >= per_run:
                        break
                    addr = (v.recruiter_email or "").lower()
                    if addr in offered_to:
                        continue
                    # Черновик уже есть (предложили раньше или есть черновик
                    # другого канала — второй схема не пустит).
                    if any(a.is_draft for a in v.applications):
                        continue
                    if self._recently_emailed(s, v.recruiter_email):
                        continue
                    text, issues = self._letter(s, v, prof_name)
                    if not text:
                        continue
                    sig = _signature()
                    body = text + (f"\n\n{sig}" if sig else "")
                    markup = {"inline_keyboard": [[
                        {"text": "📧 Отправить", "callback_data": f"emailsend:{v.id}"},
                        {"text": "⏭ Пропустить", "callback_data": f"tgskip:{v.id}"},
                    ]]}
                    msg = (f"✉️ <b>Отклик по почте</b> · {score}%\n"
                           f"{_esc(v.role or '—')} — {_esc(v.company or '—')}\n"
                           f"Кому: <code>{_esc(v.recruiter_email)}</code>\n"
                           f"Тема: {_esc(_subject(v.role or ''))}\n"
                           + (f"{_esc(v.link)}\n" if v.link else "")
                           + f"\n{_esc(body)}")
                    if _send(msg[:4000], markup) is None:
                        continue
                    now = datetime.now(timezone.utc).replace(tzinfo=None)
                    s.add(Application(vacancy_id=v.id, message_text=body, channel="email",
                                      is_draft=True, created_at=now))
                    v.status = VacancyStatus.drafted
                    s.commit()
                    offered += 1
                    offered_to.add(addr)
        finally:
            s.close()
        log.info("[email] режим %s: отправлено %d, предложено %d, в очереди %d",
                 m, sent, offered, len(q))
        return {"agent": self.name, "mode": m, "sent": sent, "offered": offered, "queue": len(q)}
