"""Живость сессии hh.ru: проверка, заблаговременное предупреждение, тревога.

Зачем
-----
Отклик на hh.ru ходит с cookies из storage_state Playwright
(HH_STATE_PATH, по умолчанию /root/autoapply/state.json). Логин проходит
человек руками на машине с экраном — на сервере дисплея нет. Когда сессия
умирает, отклики просто перестают уходить: очередь копится, прогон
отчитывается «ошибка», и узнаётся об этом, когда кто-то заглянет в журнал.
Здесь эта дыра и закрывается: каждый прогон проверяет сессию и, если она
кончилась или кончается, пишет в телеграм с инструкцией, что сделать.

Два разных вопроса, и путать их нельзя
--------------------------------------
1. КОГДА ИСТЕКАЕТ COOKIE — видно из файла, без сети, мгновенно. Но это
   только ВЕРХНЯЯ ГРАНИЦА: hhtoken hh выписывает примерно на 400 дней, и
   до этой даты дело обычно не доходит.
2. ЖИВА ЛИ СЕССИЯ НА САМОМ ДЕЛЕ — знает только hh. Сервер гасит сессию
   раньше срока cookie: смена пароля, выход на другом устройстве,
   антибот. Узнать можно единственным способом — сходить на страницу,
   которую отдают только своим, и посмотреть, не редиректит ли на логин.

Поэтому предупреждение «за пару дней» честно работает лишь для случая (1)
— когда cookie действительно доживает до своего срока. Основной же
детектор смерти — живая проба (2), и она ловит факт постфактум, в первый
же прогон после гибели сессии. Обещать предупреждение за два дня для
серверного разлогина было бы враньём: такого сигнала hh не даёт.

Проба стоит запуска Chromium, поэтому по расписанию она делается не
каждый час, а раз в HH_SESSION_PROBE_EVERY_H часов; дешёвая проверка
файла — каждый прогон.

Настройки (все через окружение):
    HH_STATE_PATH             путь к state.json
    HH_SESSION_WARN_DAYS      за сколько дней предупреждать (по умолчанию 2)
    HH_SESSION_PROBE_EVERY_H  как часто ходить живой пробой (по умолчанию 6)
    HH_SESSION_ALERT_REPEAT_H как часто повторять тревогу (по умолчанию 24)
    JOBSIGNAL_HOST            адрес сервера для инструкции в сообщении
"""
from __future__ import annotations

import enum
import hashlib
import json
import logging
import os
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from pathlib import Path

log = logging.getLogger("jobsignal")

DEFAULT_STATE_PATH = "/root/autoapply/state.json"

# Страница, которую hh отдаёт только залогиненному: гостя редиректит на
# логин. Проверяем поведением, а не вёрсткой меню — вёрстка ломалась дважды.
BASE_URL = "https://hh.ru"
LOGIN_PROBE_URL = f"{BASE_URL}/applicant/resumes"
LOGIN_URL_MARKERS = ("/account/login", "/account/signup", "/auth/applicant")

# Cookies, без которых сессия не поднимется, даже если файл непустой.
# Держать в синхроне с AUTH_COOKIES в /root/autoapply/tools/hh_login_local.py.
AUTH_COOKIES = ("hhtoken", "hhuid")

WARN_DAYS = float(os.environ.get("HH_SESSION_WARN_DAYS", "2"))
PROBE_EVERY_H = float(os.environ.get("HH_SESSION_PROBE_EVERY_H", "6"))
ALERT_REPEAT_H = float(os.environ.get("HH_SESSION_ALERT_REPEAT_H", "24"))
JOURNAL = Path(os.environ.get("HH_SESSION_JOURNAL", "data/hh_session.json"))


class Level(str, enum.Enum):
    OK = "ok"          # сессия рабочая
    WARN = "warn"      # cookie доживает последние дни
    DEAD = "dead"      # войти не получится: файла нет, побит, истёк или разлогинили


def state_path() -> Path:
    return Path(os.environ.get("HH_STATE_PATH", DEFAULT_STATE_PATH))


def _now() -> datetime:
    return datetime.now(timezone.utc)


def _iso(dt: datetime | None) -> str | None:
    return dt.isoformat() if dt else None


def _parse(value: str | None) -> datetime | None:
    if not value:
        return None
    try:
        return datetime.fromisoformat(value)
    except ValueError:
        return None


def human_delta(delta: timedelta) -> str:
    """«3 дн. 4 ч» — читаемый срок для сообщения человеку."""
    secs = int(delta.total_seconds())
    sign = "" if secs >= 0 else "минус "
    secs = abs(secs)
    days, rest = divmod(secs, 86400)
    hours, rest = divmod(rest, 3600)
    minutes = rest // 60
    if days:
        return f"{sign}{days} дн. {hours} ч"
    if hours:
        return f"{sign}{hours} ч {minutes} мин"
    return f"{sign}{minutes} мин"


def human_time(dt: datetime | None) -> str:
    return dt.astimezone().strftime("%d.%m.%Y %H:%M %Z") if dt else "—"


@dataclass
class SessionInfo:
    """Что видно по файлу сессии, без обращения к hh."""

    path: Path
    exists: bool = False
    readable: bool = True
    cookies: int = 0
    hh_cookies: int = 0
    missing: list[str] = field(default_factory=list)
    # Ближайший срок среди cookies аутентификации. None — их нет или они
    # сессионные (без срока).
    expires_at: datetime | None = None
    per_cookie: dict[str, datetime | None] = field(default_factory=dict)
    refreshed_at: datetime | None = None   # mtime файла: его переписывает удачный прогон
    fingerprint: str = ""                  # отпечаток hhtoken — меняется при новом логине

    @property
    def expires_in(self) -> timedelta | None:
        return self.expires_at - _now() if self.expires_at else None

    @property
    def cookie_expired(self) -> bool:
        return self.expires_at is not None and self.expires_at <= _now()

    @property
    def cookie_expiring(self) -> bool:
        left = self.expires_in
        return left is not None and timedelta() < left <= timedelta(days=WARN_DAYS)


def inspect_state(path: Path | None = None) -> SessionInfo:
    """Разобрать state.json. В сеть не ходит, поэтому дёшево и на каждый прогон."""
    path = Path(path) if path else state_path()
    info = SessionInfo(path=path)
    if not path.exists():
        info.missing = list(AUTH_COOKIES)
        return info
    info.exists = True
    info.refreshed_at = datetime.fromtimestamp(path.stat().st_mtime, tz=timezone.utc)
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        cookies = data["cookies"]
        if not isinstance(cookies, list):
            raise TypeError(f"cookies — {type(cookies).__name__}, ожидался список")
    except Exception as exc:  # noqa: BLE001 — битый файл сессии = сессии нет
        log.error("[hh_session] %s не читается: %s", path, exc)
        info.readable = False
        info.missing = list(AUTH_COOKIES)
        return info

    info.cookies = len(cookies)
    hh = [c for c in cookies if "hh.ru" in (c.get("domain") or "")]
    info.hh_cookies = len(hh)
    by_name = {c.get("name"): c for c in hh}
    info.missing = [n for n in AUTH_COOKIES if n not in by_name]

    expiries: list[datetime] = []
    for name in AUTH_COOKIES:
        cookie = by_name.get(name)
        stamp = None
        if cookie:
            raw = cookie.get("expires")
            # -1 или отсутствие — cookie сессионная: срока у неё нет, и это
            # не повод считать её истёкшей.
            if isinstance(raw, (int, float)) and raw > 0:
                stamp = datetime.fromtimestamp(raw, tz=timezone.utc)
                expiries.append(stamp)
        info.per_cookie[name] = stamp
    if expiries:
        info.expires_at = min(expiries)

    token = by_name.get("hhtoken", {}).get("value") or ""
    if token:
        info.fingerprint = hashlib.sha256(token.encode()).hexdigest()[:16]
    return info


# ── журнал наблюдений ────────────────────────────────────────────────────
# Нужен для трёх вещей: не слать одну и ту же тревогу каждый час, помнить
# последнюю удачную пробу и считать возраст ИМЕННО ЭТОЙ сессии. Возраст
# считается от первой встречи с отпечатком hhtoken, а не от даты логина:
# когда человек логинился, в файле не записано, и придумывать эту дату
# нельзя. Поэтому в отчёте возраст помечен как «наблюдаем с».

def read_journal() -> dict:
    try:
        return json.loads(JOURNAL.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return {}
    except Exception as exc:  # noqa: BLE001 — журнал вспомогательный
        log.warning("[hh_session] журнал %s не читается: %s", JOURNAL, exc)
        return {}


def write_journal(data: dict) -> None:
    try:
        JOURNAL.parent.mkdir(parents=True, exist_ok=True)
        JOURNAL.write_text(json.dumps(data, ensure_ascii=False, indent=2),
                           encoding="utf-8")
    except Exception as exc:  # noqa: BLE001
        log.warning("[hh_session] журнал %s не пишется: %s", JOURNAL, exc)


def _track_fingerprint(journal: dict, info: SessionInfo) -> dict:
    """Отследить смену сессии: другой hhtoken — другой логин, отсчёт заново."""
    if not info.fingerprint:
        return journal
    if journal.get("fingerprint") != info.fingerprint:
        journal["fingerprint"] = info.fingerprint
        journal["first_seen_at"] = _iso(_now())
        journal.pop("last_alive_at", None)
        journal.pop("alerts", None)
        log.info("[hh_session] новая сессия hh.ru (отпечаток %s)", info.fingerprint)
    return journal


# ── живая проба ──────────────────────────────────────────────────────────

async def check_page(page) -> tuple[bool, str]:
    """Жива ли сессия в уже открытой странице Playwright.

    Единственный источник правды о логине для всего проекта: hh_apply
    зовёт эту же функцию. Опираемся на поведение, а не на вёрстку — гостя
    hh редиректит с /applicant/resumes на /account/login.
    """
    try:
        await page.goto(LOGIN_PROBE_URL, wait_until="domcontentloaded")
    except Exception as exc:  # noqa: BLE001
        return False, f"страница {LOGIN_PROBE_URL} не открылась: {exc}"
    url = page.url
    if any(marker in url for marker in LOGIN_URL_MARKERS):
        return False, f"hh редиректит на логин ({url}) — сессия разлогинена"
    if "hh.ru" not in url:
        return False, f"неожиданный редирект на {url}"
    return True, url


async def _probe_async(path: Path, headless: bool) -> tuple[bool, str]:
    from playwright.async_api import async_playwright

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=headless)
        try:
            ctx = await browser.new_context(
                storage_state=str(path) if path.exists() else None,
                locale="ru-RU",
                viewport={"width": 1440, "height": 900},
            )
            ctx.set_default_timeout(int(os.environ.get("HH_NAV_TIMEOUT_MS", "30000")))
            page = await ctx.new_page()
            alive, detail = await check_page(page)
            # Куки продлеваются на каждом визите — сохраняем, но ТОЛЬКО если
            # сессия рабочая: иначе рабочий state.json затрётся гостевым.
            if alive:
                try:
                    await ctx.storage_state(path=str(path))
                except Exception as exc:  # noqa: BLE001
                    log.warning("[hh_session] state.json не сохранён: %s", exc)
            return alive, detail
        finally:
            try:
                await browser.close()
            except Exception as exc:  # noqa: BLE001
                log.warning("[hh_session] браузер не закрылся: %s", exc)


def probe(path: Path | None = None, headless: bool = True) -> tuple[bool, str]:
    """Сходить на hh и узнать, пускают ли нас. Запускает Chromium."""
    import asyncio

    path = Path(path) if path else state_path()
    try:
        return asyncio.run(_probe_async(path, headless))
    except Exception as exc:  # noqa: BLE001 — проба не должна ронять прогон
        log.error("[hh_session] проба не выполнилась: %s", exc, exc_info=True)
        return False, f"проба не выполнилась ({type(exc).__name__}: {exc})"


def record_probe(alive: bool) -> None:
    """Записать результат пробы, сделанной чужим кодом (отклик уже в браузере).

    Отклик открывает hh сам и проверяет логин по дороге. Если не записать
    это наблюдение, журнал будет считать сессию живой по устаревшей
    плановой пробе, а «последний раз пускали» в тревоге покажет время
    полусуточной давности вместо настоящего.
    """
    journal = _track_fingerprint(read_journal(), inspect_state())
    journal["last_probe_at"] = _iso(_now())
    if alive:
        journal["last_alive_at"] = journal["last_probe_at"]
    write_journal(journal)


# ── сводка и решение ─────────────────────────────────────────────────────

@dataclass
class SessionStatus:
    level: Level
    reason: str
    info: SessionInfo
    probed: bool = False
    probe_detail: str = ""
    first_seen_at: datetime | None = None
    last_alive_at: datetime | None = None
    last_probe_at: datetime | None = None


def _classify(info: SessionInfo) -> tuple[Level, str] | None:
    """Приговор по одному файлу, без сети. None — файл в порядке, нужна проба."""
    if not info.exists:
        return Level.DEAD, f"файла сессии нет: {info.path}"
    if not info.readable:
        return Level.DEAD, f"файл сессии не читается: {info.path}"
    if info.missing:
        return Level.DEAD, ("в файле нет cookies входа: "
                            + ", ".join(info.missing))
    if info.cookie_expired:
        return Level.DEAD, (f"cookie входа истекла "
                            f"{human_time(info.expires_at)}")
    if info.cookie_expiring:
        return Level.WARN, (f"cookie входа истекает через "
                            f"{human_delta(info.expires_in)} — "
                            f"{human_time(info.expires_at)}")
    return None


def status(probe_live: bool | None = None, headless: bool = True) -> SessionStatus:
    """Полная сводка. probe_live=None — сходить живой пробой, если пора.

    Порядок важен: сначала бесплатная проверка файла. Если войти заведомо
    нечем, Chromium ради этого не поднимаем.
    """
    info = inspect_state()
    journal = _track_fingerprint(read_journal(), info)

    verdict = _classify(info)
    st = SessionStatus(
        level=verdict[0] if verdict else Level.OK,
        reason=verdict[1] if verdict else "",
        info=info,
        first_seen_at=_parse(journal.get("first_seen_at")),
        last_alive_at=_parse(journal.get("last_alive_at")),
        last_probe_at=_parse(journal.get("last_probe_at")),
    )

    if verdict and verdict[0] is Level.DEAD:
        write_journal(journal)
        return st

    if probe_live is None:
        last = st.last_probe_at
        probe_live = last is None or (_now() - last) >= timedelta(hours=PROBE_EVERY_H)
    if not probe_live:
        if not verdict:
            st.reason = "по файлу сессия в порядке (живой пробы в этот раз не было)"
        write_journal(journal)
        return st

    alive, detail = probe(info.path, headless=headless)
    st.probed = True
    st.probe_detail = detail
    journal["last_probe_at"] = _iso(_now())
    if alive:
        journal["last_alive_at"] = journal["last_probe_at"]
        st.last_alive_at = _parse(journal["last_alive_at"])
        st.last_probe_at = st.last_alive_at
        if not verdict:
            st.reason = "hh пустил на страницу для своих — сессия рабочая"
    else:
        st.level = Level.DEAD
        st.reason = detail
        st.last_probe_at = _parse(journal["last_probe_at"])
    write_journal(journal)
    return st


# ── уведомление ──────────────────────────────────────────────────────────

def _host() -> str:
    return os.environ.get("JOBSIGNAL_HOST", "5.42.104.39")


def _howto() -> str:
    host = _host()
    return (
        "<b>Что сделать</b> (на компьютере с экраном — на сервере дисплея "
        "нет, hh требует пароль, SMS и капчу):\n"
        "1. Забрать скрипт входа:\n"
        f"<code>scp root@{host}:/root/autoapply/tools/hh_login_local.py .</code>\n"
        "2. Поставить playwright (один раз):\n"
        "<code>python3 -m venv ~/.hh &amp;&amp; ~/.hh/bin/pip install playwright "
        "&amp;&amp; ~/.hh/bin/playwright install chromium</code>\n"
        "3. Войти и залить сессию — откроется окно браузера, "
        "логинишься руками:\n"
        f"<code>~/.hh/bin/python hh_login_local.py --upload "
        f"root@{host}:/root/autoapply/state.json</code>\n"
        "4. Проверить на сервере:\n"
        "<code>cd /opt/jobsignal_local &amp;&amp; .venv/bin/python run.py "
        "hh-session --probe</code>"
    )


def _message(st: SessionStatus) -> str:
    info = st.info
    head = (
        "🔴 <b>jobsignal: сессия hh.ru не работает</b>"
        if st.level is Level.DEAD
        else "🟡 <b>jobsignal: сессия hh.ru скоро кончится</b>"
    )
    lines = [head, "", f"Причина: {st.reason}", f"Файл: <code>{info.path}</code>"]
    if info.refreshed_at:
        lines.append(f"Обновлён: {human_time(info.refreshed_at)}")
    if st.last_alive_at:
        lines.append(f"Последний раз пускали: {human_time(st.last_alive_at)} "
                     f"({human_delta(_now() - st.last_alive_at)} назад)")
    if st.level is Level.DEAD:
        lines += [
            "",
            "Пока так: отклики на hh.ru не уходят, очередь копится. "
            "Сбор вакансий и телеграм-часть работают — они без сессии.",
        ]
    else:
        lines += ["", "Пока всё работает — это запас времени, не авария."]
    lines += ["", _howto()]
    return "\n".join(lines)


def _alert_due(journal: dict, level: Level) -> bool:
    """Не повторять одну и ту же тревогу каждый час: конвейер идёт ежечасно."""
    last = _parse((journal.get("alerts") or {}).get(level.value))
    return last is None or (_now() - last) >= timedelta(hours=ALERT_REPEAT_H)


def notify(st: SessionStatus, force: bool = False) -> bool:
    """Отправить тревогу в телеграм, если она новая. True — отправили."""
    if st.level is Level.OK:
        return False
    journal = read_journal()
    if not force and not _alert_due(journal, st.level):
        log.info("[hh_session] тревога «%s» уже отправлена за последние %g ч "
                 "— не повторяю", st.level.value, ALERT_REPEAT_H)
        return False

    from .agents.notify_bot import _send

    if _send(_message(st)) is None:
        # Не отправилось — время НЕ отмечаем, чтобы следующий прогон повторил.
        log.error("[hh_session] тревогу в телеграм отправить не удалось")
        return False
    journal.setdefault("alerts", {})[st.level.value] = _iso(_now())
    write_journal(journal)
    log.info("[hh_session] тревога «%s» отправлена в телеграм", st.level.value)
    return True


def report_dead(detail: str) -> SessionStatus:
    """Сессия не прошла живую проверку в чужом коде — оформить и уведомить.

    Отдельная функция, а не status(probe_live=False): по файлу такая сессия
    выглядит целой (cookies на месте, срок не вышел), и проверка файла
    вернула бы OK. Приговор здесь выносит уже случившийся редирект на
    логин, файл нужен только для подробностей в сообщении.
    """
    journal = read_journal()
    st = SessionStatus(
        level=Level.DEAD,
        reason=detail,
        info=inspect_state(),
        probed=True,
        probe_detail=detail,
        first_seen_at=_parse(journal.get("first_seen_at")),
        last_alive_at=_parse(journal.get("last_alive_at")),
        last_probe_at=_parse(journal.get("last_probe_at")),
    )
    notify(st)
    return st


def guard(probe_live: bool | None = None) -> SessionStatus:
    """Проверка в начале прогона: посмотреть, при беде — написать в телеграм.

    Прогон не роняет и ничего не блокирует: сбор вакансий и телеграм-часть
    от сессии hh не зависят, а отклик и без того упрётся в свою проверку.
    Задача одна — чтобы человек узнал сам, а не из журнала.
    """
    st = status(probe_live=probe_live)
    if st.level is Level.OK:
        log.info("[hh_session] сессия hh.ru в порядке: %s", st.reason)
    else:
        log.error("[hh_session] сессия hh.ru — %s: %s", st.level.value, st.reason)
        notify(st)
    return st


def report(st: SessionStatus) -> str:
    """Человекочитаемая сводка для CLI."""
    info = st.info
    mark = {Level.OK: "✅ рабочая", Level.WARN: "🟡 доживает",
            Level.DEAD: "🔴 не работает"}[st.level]
    out = [
        f"Сессия hh.ru: {mark}",
        f"  {st.reason}",
        f"  файл:               {info.path}",
        f"  cookies:            {info.cookies} всего, {info.hh_cookies} для hh.ru",
    ]
    if info.missing:
        out.append(f"  НЕ ХВАТАЕТ cookies:  {', '.join(info.missing)}")
    for name, stamp in info.per_cookie.items():
        if stamp is None:
            out.append(f"  {name + ':':<20}(сессионная, без срока)")
        else:
            out.append(f"  {name + ':':<20}{human_time(stamp)} "
                       f"(осталось {human_delta(stamp - _now())})")
    if info.expires_at:
        out.append(f"  срок по cookies:    {human_time(info.expires_at)} "
                   f"— через {human_delta(info.expires_in)}")
        out.append("  (это верхняя граница: hh гасит сессию и раньше срока — "
                   "смена пароля, выход на другом устройстве, антибот)")
    out.append(f"  файл обновлён:      {human_time(info.refreshed_at)}")
    if info.fingerprint:
        out.append(f"  отпечаток hhtoken:  {info.fingerprint}")
    if st.first_seen_at:
        out.append(f"  наблюдаем с:        {human_time(st.first_seen_at)} "
                   f"({human_delta(_now() - st.first_seen_at)})")
    out.append(f"  живая проба:        "
               + (f"{'пустили' if st.level is not Level.DEAD else 'не пустили'} "
                  f"— {st.probe_detail}" if st.probed else "в этот раз не делалась"))
    if st.last_alive_at:
        out.append(f"  последний раз пускали: {human_time(st.last_alive_at)}")
    return "\n".join(out)
