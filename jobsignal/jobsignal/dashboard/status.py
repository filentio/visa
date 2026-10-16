"""Сводка «Состояние» для кабинета: работает ли система и что она сделала.

Вопросы, ради которых человек заходил на сервер по SSH, на одной странице:
жива ли сессия hh, сколько осталось трафика прокси, уходят ли отклики,
что ответили, какие таймеры упали. Каждый блок собирается отдельно: сбой
одного (нет systemctl, нет таблицы) показывается в этом блоке и не роняет
страницу целиком.
"""
from __future__ import annotations

import logging
import os
import sqlite3
import subprocess
from contextlib import contextmanager
from datetime import datetime, timezone

log = logging.getLogger("jobsignal")

def _db_path() -> str:
    """Тот же файл, что у ORM (db._default_url), а не относительный путь:
    иначе запуск из другого каталога открыл бы чужую или пустую базу."""
    from jobsignal.db import _default_url
    return _default_url().replace("sqlite:///", "", 1)

# Юнит → человеческое имя. Таймерные — с расписанием, остальные всегда в работе.
UNITS = [
    ("jobsignal-pipeline", "Часовой прогон: сбор, разбор, оценка", True),
    ("jobsignal-hh-apply", "Автоотклик hh", True),
    ("jobsignal-replies", "Учёт ответов hh", True),
    ("jobsignal-tgwatch", "Сторож входящих в Telegram", True),
    ("jobsignal-backup", "Ночная копия базы", True),
    ("jobsignal-bot", "Бот уведомлений", False),
]


def _block(fn):
    try:
        return fn()
    except Exception as exc:  # noqa: BLE001 — один блок не роняет страницу
        log.warning("[status] %s: %s", fn.__name__, exc)
        return {"error": f"{type(exc).__name__}: {exc}"}


@contextmanager
def _db():
    # with sqlite3.connect() закрывает транзакцию, но не соединение.
    c = sqlite3.connect(f"file:{_db_path()}?mode=ro", uri=True, timeout=10)
    c.row_factory = sqlite3.Row
    try:
        yield c
    finally:
        c.close()


def _iso(dt) -> str | None:
    if dt is None:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.isoformat()


def hh_session_block() -> dict:
    from jobsignal import hh_session
    # Без живой пробы: страница не должна поднимать Chromium и тратить прокси.
    st = hh_session.status(probe_live=False)
    return {
        "level": st.level.value,
        "reason": st.reason,
        "last_alive_at": _iso(st.last_alive_at),
        "expires_at": _iso(st.info.expires_at),
    }


def proxy_block() -> dict:
    from jobsignal import proxy_meter, hh_session
    if not hh_session.proxy_url():
        return {"disabled": True}
    return proxy_meter.summary()


def queue_block() -> dict:
    from jobsignal.ratelimit import rate_status
    auto = int(os.environ.get("HH_AUTO_THRESHOLD", "80"))
    with _db() as c:
        queue = c.execute(
            "select count(*) from vacancies v where contact_type='hh' "
            "and is_primary=1 and status in ('matched','drafted') "
            "and (select max(score) from match_scores where vacancy_id=v.id) >= ? "
            "and id not in (select vacancy_id from applications "
            "where sent_at is not null)", (auto,)).fetchone()[0]
        parse_queue = c.execute(
            "select count(*) from raw_posts where parsed=0 and awaiting_details=0 "
            "and text is not null and coalesce(parse_attempts,0) < 3").fetchone()[0]
    rate = rate_status()
    return {"threshold": auto, "hh_queue": queue, "parse_queue": parse_queue,
            "sent_hour": rate.get("sent_hour"), "per_hour": rate.get("per_hour"),
            "sent_day": rate.get("sent_day"), "per_day": rate.get("per_day")}


def funnel_block() -> dict:
    out = {}
    with _db() as c:
        for label, since in (("day", "-24 hours"), ("week", "-7 days")):
            sent = dict(c.execute(
                "select channel, count(*) from applications where sent_at is not null "
                "and sent_at > datetime('now', ?) group by 1", (since,)).fetchall())
            replies = dict(c.execute(
                "select coalesce(reply_state, 'REPLIED'), count(*) from applications "
                "where replied_at is not null and replied_at > datetime('now', ?) "
                "group by 1", (since,)).fetchall())
            new_vac = c.execute(
                "select count(*) from vacancies where is_primary=1 "
                "and created_at > datetime('now', ?)", (since,)).fetchone()[0]
            out[label] = {"sent": sent, "sent_total": sum(sent.values()),
                          "replies": replies, "new_vacancies": new_vac}
        total = c.execute("select count(*) from applications where sent_at is not null "
                          "and channel='hh'").fetchone()[0]
        states = dict(c.execute(
            "select coalesce(reply_state, 'NO_REPLY'), count(*) from applications "
            "where sent_at is not null and channel='hh' group by 1").fetchall())
    out["all_hh"] = {"sent": total, "states": states}
    return out


def recent_block(limit: int = 12) -> list[dict]:
    with _db() as c:
        rows = c.execute(
            "select a.sent_at, a.channel, a.reply_state, a.replied_at, "
            "v.id, v.role, v.company, v.link, "
            "(select max(score) from match_scores where vacancy_id=v.id) score "
            "from applications a join vacancies v on v.id=a.vacancy_id "
            "where a.sent_at is not null order by a.sent_at desc limit ?",
            (limit,)).fetchall()
    return [dict(r) for r in rows]


def _show(unit: str, props: str) -> dict:
    out = subprocess.run(["systemctl", "show", unit, "-p", props],
                         capture_output=True, text=True, timeout=10)
    return dict(line.split("=", 1) for line in out.stdout.splitlines() if "=" in line)


def units_block() -> list[dict]:
    res = []
    for unit, title, timer in UNITS:
        svc = _show(f"{unit}.service",
                    "ActiveState,Result,ExecMainStatus,ExecMainExitTimestamp,"
                    "ExecMainStartTimestamp")
        row = {"unit": unit, "title": title,
               "active": svc.get("ActiveState", ""),
               "result": svc.get("Result", ""),
               "exit_status": svc.get("ExecMainStatus", ""),
               "last_run": svc.get("ExecMainExitTimestamp")
               or svc.get("ExecMainStartTimestamp") or ""}
        if timer:
            t = _show(f"{unit}.timer", "ActiveState,NextElapseUSecRealtime")
            row["timer_active"] = t.get("ActiveState", "")
            row["next_run"] = t.get("NextElapseUSecRealtime", "")
        # Юнит упал, если последний запуск закончился не success. Для таймерных
        # служб ActiveState=failed держится до следующего удачного запуска.
        # Постоянная служба (бот) должна просто работать.
        if timer:
            row["ok"] = (row["result"] in ("success", "") and row["active"] != "failed"
                         and row.get("timer_active") == "active")
        else:
            row["ok"] = row["active"] == "active"
        res.append(row)
    return res


def collect() -> dict:
    return {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "hh_session": _block(hh_session_block),
        "proxy": _block(proxy_block),
        "queue": _block(queue_block),
        "funnel": _block(funnel_block),
        "recent": _block(recent_block),
        "units": _block(units_block),
    }
