"""Вход на hh.ru руками — на машине с экраном, с сохранением сессии.

Зачем отдельный скрипт. На сервере дисплея нет, а hh при входе спрашивает
пароль, присылает SMS и показывает капчу. Поэтому логин делает человек на
своём компьютере, а на сервер уезжает только файл сессии (storage_state
Playwright): cookies, по которым hh потом узнаёт нас без пароля.

ВАЖНО — про прокси. hh привязывает сессию к адресу. Если войти из дома, а
пользоваться сессией с московского адреса прокси, hh её погасит: 03.10 так и
вышло — cookies остались в файле, но hh перестал их признавать, отдавал все
страницы в гостевом виде, и отклики уходили в «поле письма не найдено».
Поэтому входить надо ЧЕРЕЗ ТОТ ЖЕ ПРОКСИ, которым ходит сервер. Скрипт берёт
его из --proxy или из переменной HH_PROXY.

Проще всего (сам возьмёт прокси с сервера и сам отправит сессию обратно):

    python hh_login_local.py --via-server root@5.187.4.252

Запуск на своём компьютере (не на сервере):

    python3 -m venv ~/.hh
    ~/.hh/bin/pip install playwright
    ~/.hh/bin/playwright install chromium

    ~/.hh/bin/python hh_login_local.py \\
        --proxy http://логин:пароль@2.59.169.68:5000 \\
        --upload root@130.17.20.195:/root/autoapply/state.json

Откроется окно браузера. Войди обычным способом. Как только появится меню
соискателя, скрипт сам сохранит сессию и (если задан --upload) отправит её
на сервер через scp — пароль от сервера спросит scp, скрипт его не видит.
"""
from __future__ import annotations

import argparse
import asyncio
import os
import subprocess
import sys
from pathlib import Path
from urllib.parse import urlsplit

LOGIN_URL = "https://hh.ru/account/login"
CHECK_URL = "https://hh.ru/applicant/resumes"

# Признаки того, что вход состоялся: меню соискателя в шапке. Проверять по
# ним, а не по адресу страницы: hh перестал перебрасывать гостя на логин —
# отдаёт тот же адрес в гостевом виде, и проверка по редиректу врёт.
OWNER_SELECTORS = ('[data-qa="mainmenu_profileAndResumes"]',
                   '[data-qa="mainmenu_applicantProfilePage"]',
                   '[data-qa="mainmenu_negotiations"]')

WAIT_MINUTES = 15


def playwright_proxy(url: str | None) -> dict | None:
    if not url:
        return None
    url = url.strip().strip('"').strip("'").lstrip("﻿")
    if url.startswith("HH_PROXY="):
        url = url.split("=", 1)[1]
    # Кабинеты прокси отдают «логин:пароль@хост:порт» без схемы; curl тогда
    # подставляет http:// сам, и строка выглядела рабочей. Делаем так же.
    if url and "://" not in url:
        url = "http://" + url
    # Разбираем руками, а не urlsplit. В пароле прокси бывают «#», «/», «?»:
    # для стандартного разбора это конец адреса, и строка ломается — так
    # 14.10 падал вход с Windows, хотя curl с той же строкой работал (он
    # разбирает проще). Здесь логин:пароль — всё до ПОСЛЕДНЕЙ «@».
    scheme, sep, rest = url.partition("://")
    creds, at, hostport = rest.rpartition("@")
    if not at:
        creds, hostport = "", rest
    host, _, port = hostport.partition(":")
    port = port.rstrip("/")
    problems = []
    if not sep or scheme.lower() not in ("http", "https", "socks5"):
        # Начало строки не показываем: без «://» там логин и пароль.
        problems.append("нет http:// в начале"
                        if not sep else f"схема {scheme[:10]!r} вместо http://")
    if not host:
        problems.append("нет хоста")
    if not port.isdigit():
        problems.append("нет порта")
    if problems:
        raise SystemExit(
            "строка прокси не похожа на адрес прокси: " + ", ".join(problems)
            + "\nНужно: http://логин:пароль@хост:порт"
            + f"\nДлина строки: {len(url)}, после @: {hostport[:40]!r}")
    out = {"server": f"{scheme.lower()}://{host}:{port}"}
    if creds:
        user, _, pwd = creds.partition(":")
        out["username"] = user
        if pwd:
            out["password"] = pwd
    return out


TUNNEL_PORT = 18765


def proxy_via_server(server: str) -> tuple[str, subprocess.Popen]:
    """Прокси сервера через SSH-туннель: строку прокси берём с сервера, сами
    подключаемся к нему с адреса сервера.

    Зачем. Строку прокси с паролем руками вставлять неудобно (в PowerShell
    ломаются кавычки и спецсимволы), а прокси принимает подключения только с
    разрешённых адресов: 14.10 домашний адрес он отверг, и спасло ровно это —
    туннель через сервер. Сессия при этом получается с того же выходного
    адреса прокси, что у сервера, и hh её не гасит.
    """
    out = subprocess.run(
        ["ssh", server, "grep -m1 '^HH_PROXY=' /opt/jobsignal_local/config/.env"],
        capture_output=True, text=True, timeout=120)
    line = (out.stdout or "").strip()
    if out.returncode or not line:
        raise SystemExit(f"не удалось прочитать HH_PROXY с {server}: "
                         f"{(out.stderr or '').strip()[:200] or 'строки нет'}")
    url = line.split("=", 1)[1].strip().strip('"').strip("'")
    if "://" not in url:
        url = "http://" + url
    scheme, _, rest = url.partition("://")
    creds, at, hostport = rest.rpartition("@")
    if not at:
        creds, hostport = "", rest
    print("Поднимаю SSH-туннель. Если спросит пароль сервера — введи его ещё раз.")
    tunnel = subprocess.Popen(
        ["ssh", "-N", "-o", "ExitOnForwardFailure=yes",
         "-L", f"{TUNNEL_PORT}:{hostport.rstrip('/')}", server])
    # Ждём, пока туннель реально начнёт принимать подключения: при входе по
    # паролю ssh висит на запросе пароля, и браузер, открытый раньше, пошёл
    # бы в пустой порт.
    import socket
    import time
    deadline = time.time() + 180
    while True:
        if tunnel.poll() is not None:
            raise SystemExit("SSH-туннель не поднялся — проверь, что ssh к серверу работает")
        try:
            with socket.create_connection(("127.0.0.1", TUNNEL_PORT), timeout=1):
                break
        except OSError:
            pass
        if time.time() > deadline:
            tunnel.terminate()
            raise SystemExit("SSH-туннель за 3 минуты не поднялся")
        time.sleep(1)
    local = f"{scheme}://{creds + '@' if creds else ''}127.0.0.1:{TUNNEL_PORT}"
    print(f"прокси сервера через туннель: 127.0.0.1:{TUNNEL_PORT} → {hostport.split(':')[0]}")
    return local, tunnel


async def run(opts) -> int:
    from playwright.async_api import async_playwright

    tunnel = None
    raw_proxy = opts.proxy or os.environ.get("HH_PROXY")
    if opts.via_server:
        raw_proxy, tunnel = proxy_via_server(opts.via_server)
        if not opts.upload:
            opts.upload = f"{opts.via_server}:/root/autoapply/state.json"
    try:
        return await _run(opts, raw_proxy)
    finally:
        if tunnel is not None:
            tunnel.terminate()


async def _run(opts, raw_proxy) -> int:
    from playwright.async_api import async_playwright

    proxy = playwright_proxy(raw_proxy)
    out = Path(opts.out).expanduser()
    print(f"прокси: {proxy['server'] if proxy else 'НЕ ЗАДАН — сессия может быть погашена'}")
    print(f"файл сессии: {out}\n")

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=False, proxy=proxy)
        ctx = await browser.new_context(locale="ru-RU",
                                        viewport={"width": 1440, "height": 900})
        page = await ctx.new_page()
        # Ошибка первой загрузки не повод закрывать окно. Через прокси и
        # туннель страница может не открыться с первого раза (14.10:
        # ERR_HTTP_RESPONSE_CODE_FAILURE), а закрытое окно не даёт ни увидеть,
        # что отвечает hh, ни обновить страницу руками.
        for attempt in range(1, 4):
            try:
                resp = await page.goto(LOGIN_URL, wait_until="domcontentloaded",
                                       timeout=90_000)
                print(f"страница входа: HTTP {resp.status if resp else '—'}")
                break
            except Exception as exc:  # noqa: BLE001
                print(f"попытка {attempt}/3: страница не открылась — "
                      f"{str(exc).splitlines()[0][:150]}")
                await asyncio.sleep(5)
        else:
            print("Страница входа так и не открылась. Окно оставляю: открой "
                  "hh.ru/account/login в нём руками или обнови страницу (F5).")

        print("Окно браузера открыто. Войди на hh обычным способом —")
        print("пароль, SMS, капча. Я жду появления меню соискателя.")
        print(f"Времени: {WAIT_MINUTES} минут. Окно не закрывай.\n")

        deadline = asyncio.get_running_loop().time() + WAIT_MINUTES * 60
        while asyncio.get_running_loop().time() < deadline:
            try:
                for sel in OWNER_SELECTORS:
                    if await page.locator(sel).count():
                        print(f"Вход выполнен (нашёл {sel}).")
                        await page.goto(CHECK_URL, wait_until="domcontentloaded")
                        await page.wait_for_timeout(5000)
                        await ctx.storage_state(path=str(out))
                        cookies = await ctx.cookies("https://hh.ru")
                        names = {c["name"] for c in cookies}
                        print(f"Сессия сохранена: {out}")
                        print(f"  cookies для hh.ru: {len(cookies)}, "
                              f"hhtoken {'есть' if 'hhtoken' in names else 'НЕТ'}, "
                              f"hhuid {'есть' if 'hhuid' in names else 'НЕТ'}")
                        if "hhtoken" not in names:
                            print("  ВНИМАНИЕ: без hhtoken сессия нерабочая — "
                                  "похоже, вход не завершён")
                        await browser.close()
                        return upload(out, opts.upload)
            except Exception:  # noqa: BLE001 — страница могла перерисоваться
                pass
            await asyncio.sleep(2)

        print(f"\nЗа {WAIT_MINUTES} минут вход не завершился. Ничего не сохранял.")
        await browser.close()
        return 1


def upload(path: Path, target: str | None) -> int:
    if not target:
        print("\n--upload не задан. Отправь файл на сервер сам:")
        print(f"  scp {path} root@СЕРВЕР:/root/autoapply/state.json")
        return 0
    print(f"\nОтправляю на {target} (пароль спросит scp)…")
    rc = subprocess.call(["scp", str(path), target])
    if rc:
        print(f"scp вернул {rc}. Отправь вручную:\n  scp {path} {target}")
        return rc
    print("Готово. Проверь на сервере:")
    print("  cd /opt/jobsignal_local && .venv/bin/python run.py hh-session --probe")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(prog="hh_login_local.py")
    ap.add_argument("--proxy", default=None,
                    help="http://логин:пароль@хост:порт — ТОТ ЖЕ, что у сервера")
    ap.add_argument("--out", default="state.json", help="куда сохранить сессию")
    ap.add_argument("--upload", default=None,
                    help="root@сервер:/root/autoapply/state.json")
    ap.add_argument("--via-server", default=None, metavar="root@СЕРВЕР",
                    help="взять прокси с сервера и ходить через SSH-туннель к нему; "
                         "сессия сама уедет на сервер. Самый простой способ")
    return asyncio.run(run(ap.parse_args()))


if __name__ == "__main__":
    sys.exit(main())
