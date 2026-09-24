"""Учёт ответов на отклики hh.ru — нулевая задача docs/ROADMAP.md.

До этого агента `Application.replied_at` проставлялся только вручную двумя
кнопками в дашборде, а автоматически — никем. Поэтому «ноль ответов на сто
откликов» означало не «не отвечают», а «не смотрим»: разведка 21.09 показала
196 откликов, из них 57 отказов и 6 приглашений на собеседование. Треть
откликов получила ответ, и система об этом не знала.

Как устроено. hh отдаёт список откликов в том же блоке состояния страницы,
что и поиск вакансий (`HH-Lux-InitialState`), в ключе
`applicantNegotiations.topicList`. У каждой записи есть:

  vacancyId    — чем связываем с нашей базой (в Vacancy.link лежит тот же id)
  initialState — всегда RESPONSE: это наш отклик
  lastState    — RESPONSE (молчание) | DISCARD | INVITATION | INTERVIEW | HIRED
  lastModified — когда состояние менялось, это и есть время ответа

Признак ответа получается прямой, без эвристик: `lastState != RESPONSE`.

Отказ считается ответом намеренно. Для аналитики «прочитали и отказали» и
«не отреагировали» — разные исходы, а сейчас они оба выглядят как ноль.
Поэтому кроме `replied_at` пишем `reply_state` — что именно ответили.

Chromium здесь не нужен: страница читается обычным requests с cookies из
state.json. При 660 МБ на браузер и ~900 свободных это существенно.
"""
from __future__ import annotations

import html as htmlmod
import json
import logging
import os
import re
import time
from datetime import datetime
from pathlib import Path

import requests
from sqlalchemy import text as sqltext

from jobsignal.db import Application, Vacancy, VacancyStatus, get_session_factory

log = logging.getLogger("jobsignal")

NEGOTIATIONS_URL = "https://hh.ru/applicant/negotiations"
STATE_RE = re.compile(r'id="HH-Lux-InitialState"[^>]*>(.*?)</template>', re.S)
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ru-RU,ru;q=0.9",
}

# Потолок страниц за прогон. На 20 записей на страницу 15 страниц это 300
# откликов — с запасом к нынешним 196. Потолок нужен не ради экономии, а
# чтобы сломанная постраничность не крутила запросы бесконечно.
MAX_PAGES = int(os.environ.get("HH_REPLIES_MAX_PAGES", "15"))

# Состояние, означающее «мы написали, ответа нет». Всё остальное — ответ.
STATE_SILENT = "RESPONSE"

# Во что переводить вакансию. Ключ — lastState с hh, значение — наш статус.
# Список закрытый: незнакомое состояние не молчим, а сообщаем (см. ниже) —
# hh может завести новое, и узнать об этом лучше сразу.
STATE_TO_STATUS = {
    "DISCARD": VacancyStatus.rejected,
    "INVITATION": VacancyStatus.replied,
    "INTERVIEW": VacancyStatus.interview,
    "HIRED": VacancyStatus.offer,
}


# Пауза между страницами. Без неё десять страниц уходили одним залпом за
# десять секунд, и ddos-guard отвечал 403 — ночной прогон 22.09 упал именно
# так. Сборщик с hh держит те же 1,5 секунды (HH_PAGE_DELAY) и работает годами.
PAGE_DELAY = float(os.environ.get("HH_REPLIES_PAGE_DELAY", "1.5"))

# Сколько раз переждать блокировку внутри прогона и сколько ждать. Ограничение
# снимается быстро, поэтому дешевле подождать, чем будить человека тревогой.
BLOCK_RETRIES = int(os.environ.get("HH_REPLIES_BLOCK_RETRIES", "3"))
BLOCK_BACKOFF = (30, 90, 180)


class HHBlocked(RuntimeError):
    """hh ограничил обращения с нашего адреса (403/429).

    Отдельно от HHRepliesBroken намеренно: правило из docs/SERVICES.md —
    «отказ внешней системы различается по смыслу: лимит, блокировка,
    изменение структуры, исчезнувший объект; одинаково выглядящие отказы
    лечатся по-разному». Блокировка лечится ожиданием, смена структуры —
    правкой кода. Свалив их в одно, я получил тревогу с неверным диагнозом.
    """


class HHRepliesBroken(RuntimeError):
    """Структура страницы откликов изменилась либо сессия мертва.

    Отдельный класс, а не общее исключение: молчаливый ноль здесь — та же
    поломка, что дважды прятала от нас месяцы простоя. Пустой список при
    непустом счётчике это не «ответов нет», а «разбор не работает».
    """


def _state_path() -> Path:
    from jobsignal import hh_session
    return hh_session.state_path()


def _cookies() -> dict:
    path = _state_path()
    if not path.exists():
        raise HHRepliesBroken(f"нет файла сессии hh.ru: {path}")
    data = json.loads(path.read_text())
    jar = {c["name"]: c["value"] for c in data.get("cookies", [])
           if "hh.ru" in (c.get("domain") or "")}
    if not jar:
        raise HHRepliesBroken(f"в {path} нет cookies для hh.ru")
    return jar


def _fetch_page(cookies: dict, page: int) -> dict:
    from jobsignal import hh_session
    r = requests.get(NEGOTIATIONS_URL, params={"page": page},
                     headers=HEADERS, cookies=cookies,
                     proxies=hh_session._requests_proxies(), timeout=30)
    if "account/login" in r.url or "auth" in r.url:
        raise HHRepliesBroken(
            "hh.ru перекинул на вход: сессия протухла, нужен новый state.json")
    if r.status_code in (403, 429):
        raise HHBlocked(
            f"страница {page}: HTTP {r.status_code} — hh ограничил обращения "
            f"с нашего адреса")
    r.raise_for_status()
    m = STATE_RE.search(r.text)
    if not m:
        raise HHRepliesBroken(
            f"страница {page}: блок HH-Lux-InitialState не найден "
            f"({len(r.text)} симв.) — hh сменил разметку")
    try:
        # unescape обязателен: hh отдаёт состояние с &quot; вместо кавычек.
        state = json.loads(htmlmod.unescape(m.group(1).strip()))
    except json.JSONDecodeError as exc:
        raise HHRepliesBroken(
            f"страница {page}: состояние не разбирается как JSON: {exc}") from exc
    neg = state.get("applicantNegotiations")
    if not isinstance(neg, dict):
        raise HHRepliesBroken(
            f"страница {page}: нет ключа applicantNegotiations — структура другая")
    return {"negotiations": neg,
            "counters": state.get("applicantNegotiationsCounters") or {}}


def _fetch_page_patient(cookies: dict, page: int) -> dict:
    """_fetch_page с паузой и ожиданием блокировки.

    Блокировку пережидаем внутри прогона: она снимается за минуты, а прогон
    раз в сутки, так что ждать дешевле, чем терять сутки учёта. Если не
    отпустило за все попытки — отдаём HHBlocked наружу, и это уже повод
    сказать человеку.
    """
    if page:
        time.sleep(PAGE_DELAY)
    for attempt in range(BLOCK_RETRIES + 1):
        try:
            return _fetch_page(cookies, page)
        except HHBlocked as exc:
            if attempt >= BLOCK_RETRIES:
                raise
            wait = BLOCK_BACKOFF[min(attempt, len(BLOCK_BACKOFF) - 1)]
            log.warning("[hh_replies] %s — жду %d с (попытка %d из %d)",
                        exc, wait, attempt + 1, BLOCK_RETRIES)
            time.sleep(wait)
    raise AssertionError("недостижимо")  # pragma: no cover


def _parse_time(value: str | None) -> datetime | None:
    """'2026-09-21T22:32:15.222+03:00' → datetime. Пустое и битое — None."""
    if not value:
        return None
    try:
        return datetime.fromisoformat(value)
    except ValueError:
        log.warning("[hh_replies] не разобрал время %r", value)
        return None


def _ensure_columns(session) -> None:
    """Добавляет applications.reply_state, если его ещё нет.

    create_all() новые столбцы в существующую таблицу не добавляет — в проекте
    это уже решалось так же (notified_at у vacancies, detail_attempts у
    raw_posts). Вызов идемпотентный.
    """
    cols = {r[1] for r in session.execute(sqltext("PRAGMA table_info(applications)"))}
    if "reply_state" not in cols:
        session.execute(sqltext(
            "ALTER TABLE applications ADD COLUMN reply_state VARCHAR(32)"))
        session.commit()
        log.info("[hh_replies] добавлен столбец applications.reply_state")


def _hh_id(link: str | None) -> str:
    m = re.search(r"/vacancy/(\d+)", link or "")
    return m.group(1) if m else ""


class HHRepliesAgent:
    """Обходит свои отклики на hh.ru и проставляет факт и вид ответа."""

    name = "hh_replies"

    def __init__(self):
        self._sf = get_session_factory()

    def run(self) -> dict:
        cookies = _cookies()

        # Сначала спросить, пускают ли вообще. Пережидание блокировки внутри
        # прогона (три попытки, до пяти минут) рассчитано на короткий лимит;
        # когда закрыт весь раздел, ждать нечего — 22.09 блокировка держалась
        # часами, и прогон впустую висел бы каждую ночь.
        from jobsignal import hh_session
        blocked, why = hh_session.section_blocked()
        if blocked:
            raise HHBlocked(why)

        session = self._sf()
        _ensure_columns(session)

        # Наши отправленные отклики, разложенные по номеру вакансии на hh.
        # Одна и та же вакансия лежит в базе несколькими строками (её приносят
        # разные каналы), поэтому ключом может быть только номер hh.
        by_hh: dict[str, list[Application]] = {}
        rows = (session.query(Application, Vacancy)
                .join(Vacancy, Application.vacancy_id == Vacancy.id)
                .filter(Application.sent_at.isnot(None))
                .all())
        for app_rec, vac in rows:
            key = _hh_id(vac.link)
            if key:
                by_hh.setdefault(key, []).append(app_rec)
        log.info("[hh_replies] наших отправленных откликов с номером hh: %d",
                 len(by_hh))

        seen = matched = replied_new = unknown_state = silent = 0
        states: dict[str, int] = {}
        counters = {}
        page = 0
        total_expected = None

        while page < MAX_PAGES:
            data = _fetch_page_patient(cookies, page)
            neg = data["negotiations"]
            if page == 0:
                counters = data["counters"].get("total") or {}
                total_expected = neg.get("total")
                log.info("[hh_replies] на hh откликов: %s, счётчики: %s",
                         total_expected, counters)
            topics = neg.get("topicList") or []
            if not topics:
                if page == 0 and (total_expected or 0) > 0:
                    raise HHRepliesBroken(
                        f"страница 0: список пуст, хотя откликов {total_expected} "
                        f"— разбор не работает")
                break

            for t in topics:
                seen += 1
                last = t.get("lastState")
                states[last] = states.get(last, 0) + 1
                key = str(t.get("vacancyId") or "")
                apps = by_hh.get(key)
                if not apps:
                    # Отклик есть на hh, но у нас его нет: отправлен руками
                    # с сайта, либо вакансия пришла каналом без ссылки на hh.
                    continue
                matched += 1
                if last == STATE_SILENT:
                    # Отклик на hh есть и работодатель молчит. Это ЗНАНИЕ, и
                    # его надо записать: без отметки «подтверждённое молчание»
                    # неотличимо от «мы этот отклик вообще не видели», а таких
                    # у нас 39 из 113 (hh не показывает удалённые). В первом
                    # же разборе аналитики это дало долю ответов 5% у группы
                    # 90+ против 19% у 70-79 — при том, что знаменатель у
                    # первой состоял в основном из неизвестного.
                    for app_rec in apps:
                        app_rec.reply_state = STATE_SILENT
                    silent += 1
                    continue
                status = STATE_TO_STATUS.get(last)
                if status is None:
                    unknown_state += 1
                    log.warning("[hh_replies] незнакомое состояние %r у вакансии "
                                "hh:%s — не знаю, как трактовать", last, key)
                    continue
                when = _parse_time(t.get("lastModified"))
                for app_rec in apps:
                    if app_rec.replied_at is None:
                        app_rec.replied_at = when
                        replied_new += 1
                    # reply_state обновляем всегда: отказ мог смениться
                    # приглашением, и последнее состояние важнее первого.
                    app_rec.reply_state = last
                    vac = app_rec.vacancy
                    if vac is not None:
                        vac.status = status
            session.commit()

            pages = int(neg.get("pageCount") or 1)
            page += 1
            if page >= pages:
                break

        session.close()

        summary = {
            "agent": self.name,
            "seen": seen,
            "matched": matched,
            "replied_new": replied_new,
            "silent": silent,
            "unknown_state": unknown_state,
            "states": states,
            "counters": counters,
        }
        log.info("[hh_replies] просмотрено %d, сопоставлено с базой %d, "
                 "новых ответов %d, подтверждённого молчания %d, состояния: %s",
                 seen, matched, replied_new, silent, states)
        if seen and not matched:
            # Отклики на hh есть, но ни один не нашёлся у нас — значит
            # связывание по номеру вакансии сломано. Это поломка, а не ноль.
            raise HHRepliesBroken(
                f"просмотрено {seen} откликов, но ни один не сопоставился с "
                f"базой — связь по vacancyId не работает")
        return summary
