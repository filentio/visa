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
    p = urlsplit(url)
    out = {"server": f"{p.scheme}://{p.hostname}" + (f":{p.port}" if p.port else "")}
    if p.username:
        out["username"] = p.username
    if p.password:
        out["password"] = p.password
    return out


async def run(opts) -> int:
    from playwright.async_api import async_playwright

    proxy = playwright_proxy(opts.proxy or os.environ.get("HH_PROXY"))
    out = Path(opts.out).expanduser()
    print(f"прокси: {proxy['server'] if proxy else 'НЕ ЗАДАН — сессия может быть погашена'}")
    print(f"файл сессии: {out}\n")

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=False, proxy=proxy)
        ctx = await browser.new_context(locale="ru-RU",
                                        viewport={"width": 1440, "height": 900})
        page = await ctx.new_page()
        await page.goto(LOGIN_URL, wait_until="domcontentloaded")

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
    return asyncio.run(run(ap.parse_args()))


if __name__ == "__main__":
    sys.exit(main())
