"""Разведка: подбор вакансий, который hh делает сам под резюме.

Зачем. Замер 21.09: вакансий с баллом от 80 в очереди 391, а отправить можно
7. Ограничение системы — не оценка и не квота, а канал: автоотклик работает
только через hh, а hh даёт около 300 вакансий в неделю против 1261 из всех
источников. Из трёх способов расширить пул этот единственный не требует ни
рисковать телеграм-аккаунтом, ни учиться заполнять чужие формы: тот же канал,
который уже работает целиком — сбор, оценка, письмо, отправка, учёт ответов.

Что проверяем. Где лежит подбор и в какой форме. Если он отдаётся тем же
`vacancySearchResult.vacancies`, что и обычный поиск, то разбор карточек
(`hh_collector._card_to_item`) переиспользуется как есть.

Почему через браузер. 03.10 выяснилось, что ddos-guard на разделе кабинета
требует проверку на JavaScript: requests и подмена TLS-отпечатка получают 403
с заглушкой, настоящий Chromium через резидентный прокси проходит. Подбор под
резюме — страница для своих, значит правило то же.

Скрипт только читает. Запуск:

    cd /opt/jobsignal_local && .venv/bin/python tools/probe_hh_resume.py
"""
from __future__ import annotations

import asyncio
import html as htmlmod
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from dotenv import load_dotenv  # noqa: E402
load_dotenv("config/.env")

STATE_RE = re.compile(r'id="HH-Lux-InitialState"[^>]*>(.*?)</template>', re.S)
SETTLE_MS = 8000

# Ключи карточки, нужные разбору из hh_collector._card_to_item. Все на месте —
# подбор читается существующим кодом без изменений.
CARD_KEYS = ("vacancyId", "name", "company", "area", "compensation")


def state_of(text: str) -> dict:
    m = STATE_RE.search(text)
    if not m:
        return {}
    raw = m.group(1).strip()
    for candidate in (raw, htmlmod.unescape(raw)):
        try:
            return json.loads(candidate)
        except json.JSONDecodeError:
            continue
    return {}


def report(label: str, url: str, state: dict) -> int:
    print(f"\n=== {label}")
    print(f"    {url}")
    if not state:
        print("    состояние страницы не разобралось")
        return 0
    res = state.get("vacancySearchResult")
    if isinstance(res, dict) and isinstance(res.get("vacancies"), list):
        cards = res["vacancies"]
        print(f"    vacancySearchResult.vacancies: {len(cards)} карточек, "
              f"всего найдено: {res.get('resultsFound')}")
        if cards:
            have = [k for k in CARD_KEYS if k in cards[0]]
            missing = [k for k in CARD_KEYS if k not in cards[0]]
            print(f"    ключи для _card_to_item: {len(have)} из {len(CARD_KEYS)}")
            if missing:
                print(f"    НЕТ: {', '.join(missing)}")
            else:
                print("    разбор карточек переиспользуется как есть")
        return len(cards)
    likely = [k for k in state
              if any(w in k.lower() for w in ("vacanc", "suitable", "similar",
                                              "recommend"))]
    print(f"    vacancySearchResult.vacancies нет; похожие ключи: {likely or '—'}")
    for k in likely:
        v = state[k]
        if isinstance(v, dict):
            print(f"      {k}: объект, ключи {sorted(v)[:10]}")
        elif isinstance(v, list):
            print(f"      {k}: список, {len(v)}")
    return 0


async def main() -> int:
    from playwright.async_api import async_playwright
    from jobsignal import hh_session

    path = hh_session.state_path()
    proxy = hh_session.playwright_proxy()
    print(f"сессия: {path} ({'есть' if path.exists() else 'НЕТ'})")
    print(f"прокси: {'задан' if proxy else 'НЕ ЗАДАН — кабинет будет закрыт'}")

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=True, proxy=proxy)
        ctx = await browser.new_context(
            storage_state=str(path) if path.exists() else None,
            locale="ru-RU", viewport={"width": 1440, "height": 900})
        ctx.set_default_timeout(60000)
        page = await ctx.new_page()

        async def visit(url: str) -> dict:
            await page.goto(url, wait_until="domcontentloaded")
            await page.wait_for_timeout(SETTLE_MS)
            return state_of(await page.content())

        # Идентификаторы резюме — со страницы резюме.
        st = await visit("https://hh.ru/applicant/resumes")
        ids: list[str] = []

        def walk(node):
            if isinstance(node, dict):
                for k, v in node.items():
                    if k in ("resumeId", "hash", "_attributes") and isinstance(v, (str, int)):
                        sv = str(v)
                        if len(sv) > 5 and sv not in ids:
                            ids.append(sv)
                    walk(v)
            elif isinstance(node, list):
                for x in node[:50]:
                    walk(x)

        walk(st)
        print(f"\nидентификаторы резюме: {ids[:5] or 'не нашлись'}")
        rid = ids[0] if ids else None

        best = 0
        checks = [
            ("лента подходящих вакансий",
             "https://hh.ru/applicant/vacancy_feed"),
            ("поиск с привязкой к резюме",
             f"https://hh.ru/search/vacancy?resume={rid}&from=resumelist&items_on_page=50"
             if rid else None),
            ("похожие вакансии под резюме",
             f"https://hh.ru/applicant/similar-vacancies/{rid}" if rid else None),
        ]
        for label, url in checks:
            if not url:
                print(f"\n=== {label}\n    пропущено: нет идентификатора резюме")
                continue
            best = max(best, report(label, url, await visit(url)))

        await browser.close()

    print("\n" + "-" * 60)
    if best:
        print(f"Подбор читается, карточек за запрос: до {best}.")
    else:
        print("Ни один адрес не отдал список вакансий — нужен другой путь.")
    return 0


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
