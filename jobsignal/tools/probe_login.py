"""Разведка: действительно ли мы залогинены на hh — и где именно.

Повод. 03.10 проба сессии сказала «hh пустил на страницу для своих», а через
минуту страница вакансии открылась гостем: в шапке «Войти» и «Создать
резюме», вместо формы отклика — «Напишите телефон, чтобы работодатель мог
связаться». Адрес прокси при этом стабилен (пять замеров подряд — один и тот
же), контекст браузера один, cookies те же.

Подозрение на ложноположительную проверку. `hh_session.check_page` считает
сессию живой, если hh НЕ перебросил на /account/login. Но при проверке
ddos-guard адрес страницы не меняется: переброса нет, и проба рапортует
«живая» независимо от того, что на странице.

Скрипт смотрит не на адрес, а на содержимое: есть ли в шапке кнопка входа и
есть ли признаки личного кабинета. Ничего не отправляет и не меняет.

    cd /opt/jobsignal_local && .venv/bin/python tools/probe_login.py [URL вакансии]
"""
from __future__ import annotations

import asyncio
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from dotenv import load_dotenv  # noqa: E402
load_dotenv("config/.env")

# Признаки гостя: кнопки входа и регистрации в шапке.
GUEST_MARKS = ('[data-qa="login"]', 'text="Войти"', 'text="Создать резюме"')
# Признаки своего: меню профиля соискателя.
OWNER_MARKS = ('[data-qa="mainmenu_profileAndResumes"]',
               '[data-qa="mainmenu_applicantProfilePage"]',
               '[data-qa="mainmenu_negotiations"]')
# Признак заглушки ddos-guard.
STUB_MARKS = ("ddos-guard", "проверка браузера", "checking your browser")


async def main() -> int:
    from playwright.async_api import async_playwright
    from jobsignal import hh_session

    vacancy = sys.argv[1] if len(sys.argv) > 1 else "https://hh.ru/vacancy/137679511"
    path = hh_session.state_path()
    proxy = hh_session.playwright_proxy()
    print(f"сессия: {path} ({'есть' if path.exists() else 'НЕТ'})")
    print(f"прокси: {'задан' if proxy else 'не задан'}\n")

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=True, proxy=proxy)
        ctx = await browser.new_context(
            storage_state=str(path) if path.exists() else None,
            locale="ru-RU", viewport={"width": 1440, "height": 900})
        ctx.set_default_timeout(60000)
        page = await ctx.new_page()

        for url in ("https://hh.ru/applicant/resumes",
                    "https://hh.ru/applicant/negotiations",
                    vacancy):
            print(f"=== {url}")
            try:
                await page.goto(url, wait_until="domcontentloaded")
                # Дать ddos-guard выполнить проверку и подменить страницу.
                await asyncio.sleep(8)
            except Exception as exc:  # noqa: BLE001
                print(f"    не открылась: {exc}\n")
                continue
            body = (await page.content()).lower()
            guest = [m for m in GUEST_MARKS
                     if await page.locator(m).count()]
            owner = [m for m in OWNER_MARKS
                     if await page.locator(m).count()]
            stub = [m for m in STUB_MARKS if m in body]
            print(f"    адрес после загрузки: {page.url}")
            print(f"    размер: {len(body)} симв.")
            print(f"    признаки ГОСТЯ:  {guest or '—'}")
            print(f"    признаки СВОЕГО: {owner or '—'}")
            print(f"    заглушка:        {stub or '—'}")
            cookies = await ctx.cookies("https://hh.ru")
            names = {c["name"] for c in cookies}
            print(f"    cookies сейчас: {len(cookies)}, "
                  f"hhtoken {'есть' if 'hhtoken' in names else 'НЕТ'}, "
                  f"hhuid {'есть' if 'hhuid' in names else 'НЕТ'}\n")

        await browser.close()
    return 0


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
