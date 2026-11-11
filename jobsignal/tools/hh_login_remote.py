"""Вход на hh.ru из кабинета: браузер на сервере, управление со страницы /hh-login.

Зачем. До 08.11 сессию hh получали на своём компьютере: скрипт
tools/hh_login_local.py, venv, Playwright, SSH-туннель к прокси и scp на
сервер. Теперь браузер запускается на сервере — через тот же прокси, что у
автоотклика (hh привязывает сессию к адресу), — а человек видит его экран в
кабинете и кликает/печатает там же.

Как устроено (файлы в data/hh_login/):
  screen.jpg     снимок экрана, обновляется ~раз в 0,7 с
  status.json    {state, message, at, url}: starting → login → saving → done
                 | failed | stopped | timeout
  commands.jsonl команды от кабинета, по одной JSON-строке:
                 {"t":"click","x":..,"y":..} {"t":"type","text":".."}
                 {"t":"key","key":"Enter"} {"t":"scroll","dy":400}
                 {"t":"reload"} {"t":"login"} {"t":"stop"}

Вход замечаем по меню соискателя (как tools/hh_login_local.py). Тогда:
страница кабинета hh, проверка cookie hhtoken, под замком hh старый
state.json — в копию state.json.bak-<время>, новый — на место, запись в
журнал сессии и сообщение в телеграм.

Запускает кабинет (systemd-run, юнит jobsignal-hh-login); руками:
    cd /opt/jobsignal_local && .venv/bin/python tools/hh_login_remote.py
"""
from __future__ import annotations

import fcntl
import json
import os
import shutil
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from dotenv import load_dotenv  # noqa: E402
load_dotenv("config/.env")

DIR = Path("data/hh_login")
LOGIN_URL = os.environ.get("HH_LOGIN_URL", "https://hh.ru/account/login")
CHECK_URL = os.environ.get("HH_LOGIN_CHECK_URL", "https://hh.ru/applicant/resumes")
OWNER_SELECTORS = ('[data-qa="mainmenu_profileAndResumes"]',
                   '[data-qa="mainmenu_applicantProfilePage"]',
                   '[data-qa="mainmenu_negotiations"]')
VIEWPORT = {"width": 1280, "height": 800}
WAIT_MINUTES = int(os.environ.get("HH_LOGIN_WAIT_MIN", "15"))
FRAME_S = 0.7


def status(state: str, message: str = "", url: str = "") -> None:
    DIR.mkdir(parents=True, exist_ok=True)
    tmp = DIR / "status.tmp"
    tmp.write_text(json.dumps({"state": state, "message": message, "url": url,
                               "at": time.time()}, ensure_ascii=False))
    tmp.replace(DIR / "status.json")


def commands(offset: int) -> tuple[list[dict], int]:
    path = DIR / "commands.jsonl"
    if not path.exists():
        return [], offset
    with open(path, encoding="utf-8") as fh:
        fh.seek(offset)
        data = fh.read()
        offset = fh.tell()
    out = []
    for line in data.splitlines():
        try:
            out.append(json.loads(line))
        except ValueError:
            continue
    return out, offset


def notify(text: str) -> None:
    try:
        from jobsignal.agents.notify_bot import _send
        _send(text)
    except Exception:  # noqa: BLE001
        pass


def save_session(ctx, page) -> str:
    """Сохранить сессию под замком hh. Возвращает описание итога."""
    from jobsignal import hh_session
    status("saving", "вход выполнен — проверяю и сохраняю сессию")
    page.goto(CHECK_URL, wait_until="domcontentloaded", timeout=90_000)
    page.wait_for_timeout(4000)
    names = {c["name"] for c in ctx.cookies(CHECK_URL)}
    if "hhtoken" not in names:
        raise RuntimeError("после входа нет cookie hhtoken — вход не завершён")
    target = hh_session.state_path()
    target.parent.mkdir(parents=True, exist_ok=True)
    tmp = DIR / "state.new.json"
    ctx.storage_state(path=str(tmp))
    lock_path = os.environ.get("HH_LOCK_PATH", "/run/jobsignal-hh.lock")
    with open(lock_path, "w") as lock:
        # Ждём, пока отклик или сбор закончат с hh: подменять файл сессии под
        # работающим браузером нельзя — он перезапишет его своим.
        deadline = time.time() + 600
        while True:
            try:
                fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
                break
            except BlockingIOError:
                if time.time() > deadline:
                    raise RuntimeError("замок hh занят 10 минут — сессия не сохранена")
                status("saving", "жду, пока закончится текущая работа с hh…")
                time.sleep(5)
        if target.exists():
            shutil.copy2(target, target.with_name(
                f"{target.name}.bak-{time.strftime('%Y%m%d-%H%M%S')}"))
        shutil.move(str(tmp), str(target))
        os.chmod(target, 0o600)
    hh_session.record_probe(True)
    return f"сессия сохранена: {target} (cookies hh: {len(names)})"


def main() -> int:
    from playwright.sync_api import sync_playwright
    from jobsignal import hh_session, proxy_meter

    DIR.mkdir(parents=True, exist_ok=True)
    (DIR / "commands.jsonl").write_text("")
    status("starting", "запускаю браузер через прокси сервера")
    offset = 0
    deadline = time.time() + WAIT_MINUTES * 60
    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            headless=True, proxy=hh_session.playwright_proxy(),
            executable_path=os.environ.get("HH_LOGIN_CHROMIUM") or None)
        ctx = browser.new_context(locale="ru-RU", viewport=VIEWPORT)
        # Учёт трафика прокси — как у остальных обращений к hh.
        try:
            ctx.on("requestfinished", lambda r: proxy_meter.add(
                "hh-login", sum((r.sizes() or {}).get(k, 0) for k in
                                ("requestHeadersSize", "requestBodySize",
                                 "responseHeadersSize", "responseBodySize"))))
        except Exception:  # noqa: BLE001
            pass
        page = ctx.new_page()
        page.set_default_timeout(60_000)
        try:
            for attempt in range(1, 3):
                try:
                    page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=45_000)
                    break
                except Exception as exc:  # noqa: BLE001
                    status("starting", f"страница входа не открылась ({attempt}/2): "
                                       f"{str(exc).splitlines()[0][:120]}")
                    time.sleep(3)
            else:
                # Не открылась ни разу — почти всегда прокси (кончился трафик или
                # он недоступен). Висеть на «Запускаю браузер» бессмысленно.
                status("failed", "hh не открылся через прокси сервера — скорее всего, "
                                 "кончился трафик пакета или прокси недоступен. Проверьте "
                                 "пакет у провайдера прокси и начните вход заново.")
                return 1
            status("login", "войдите на hh: кликайте по экрану, вводите текст в поле ниже",
                   page.url)
            last_check = 0.0
            while time.time() < deadline:
                cmds, offset = commands(offset)
                for c in cmds:
                    t = c.get("t")
                    try:
                        if t == "click":
                            page.mouse.click(float(c["x"]), float(c["y"]))
                        elif t == "type":
                            page.keyboard.type(str(c.get("text", ""))[:200], delay=40)
                        elif t == "key" and c.get("key") in (
                                "Enter", "Backspace", "Tab", "Escape", "ArrowDown",
                                "ArrowUp", "Delete"):
                            page.keyboard.press(c["key"])
                        elif t == "scroll":
                            page.mouse.wheel(0, max(-2000, min(2000, int(c.get("dy", 0)))))
                        elif t == "reload":
                            page.reload(wait_until="domcontentloaded")
                        elif t == "login":
                            page.goto(LOGIN_URL, wait_until="domcontentloaded")
                        elif t == "stop":
                            status("stopped", "вход отменён, сессия не менялась")
                            return 0
                    except Exception as exc:  # noqa: BLE001 — один промах не рушит вход
                        status("login", f"команда {t} не прошла: {str(exc)[:120]}", page.url)
                try:
                    shot = DIR / "screen.tmp.jpg"
                    page.screenshot(path=str(shot), type="jpeg", quality=55)
                    shot.replace(DIR / "screen.jpg")
                except Exception:  # noqa: BLE001 — страница перерисовывается
                    pass
                if time.time() - last_check > 2:
                    last_check = time.time()
                    try:
                        if any(page.locator(s).count() for s in OWNER_SELECTORS):
                            msg = save_session(ctx, page)
                            status("done", msg + ". Автоотклик подхватит её сам.")
                            notify("✅ <b>Сессия hh.ru обновлена</b> через кабинет. "
                                   "Автоотклик и учёт ответов работают с ней.")
                            return 0
                    except RuntimeError as exc:
                        status("failed", str(exc), page.url)
                        return 1
                    except Exception:  # noqa: BLE001
                        pass
                    status("login", "войдите на hh: кликайте по экрану, вводите текст "
                                    "в поле ниже", page.url)
                time.sleep(FRAME_S)
            status("timeout", f"за {WAIT_MINUTES} минут вход не завершился — "
                              "сессия не менялась")
            return 1
        finally:
            try:
                proxy_meter.flush()
            except Exception:  # noqa: BLE001
                pass
            browser.close()


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:  # noqa: BLE001
        status("failed", f"{type(exc).__name__}: {str(exc)[:200]}")
        raise
