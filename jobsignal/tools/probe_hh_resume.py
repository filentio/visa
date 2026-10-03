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

        def keys_of(state: dict, label: str, url: str) -> None:
            """Что вообще есть в состоянии страницы — без догадок об именах."""
            print(f"\n=== {label}\n    {url}")
            if not state:
                print("    состояние не разобралось")
                return
            print(f"    ключей верхнего уровня: {len(state)}")
            interesting = [k for k in sorted(state)
                           if any(w in k.lower() for w in
                                  ("resume", "vacanc", "feed", "suitable",
                                   "similar", "recommend", "search"))]
            for k in interesting:
                v = state[k]
                size = f"[{len(v)}]" if isinstance(v, (list, dict)) else ""
                print(f"      {k}: {type(v).__name__}{size}")
                if isinstance(v, dict) and len(v) <= 20:
                    print(f"         ключи: {sorted(v)}")
            if not interesting:
                print(f"      ничего похожего; все ключи: {sorted(state)[:40]}")

        # Идентификаторы резюме ищем и в состоянии, и прямо в разметке:
        # имена ключей заранее неизвестны, а ссылка /resume/<hash> на странице
        # резюме есть почти наверняка.
        for url in ("https://hh.ru/applicant/resumes",
                    "https://hh.ru/applicant/profile/me"):
            st = await visit(url)
            keys_of(st, "страница резюме", page.url)
            body = await page.content()
            found = sorted(set(re.findall(r"/resume/([0-9a-f]{20,})", body)))
            nums = sorted(set(re.findall(r'"resumeId"\s*:\s*(\d{6,})', body)))
            print(f"    ссылки /resume/<hash>: {found[:3] or '—'}")
            print(f"    числовые resumeId:     {nums[:3] or '—'}")
            if found or nums:
                break

        ids = (found or []) + (nums or [])
        rid = ids[0] if ids else None
        print(f"\nбудем пробовать с идентификатором: {rid or 'нет'}")

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
            st = await visit(url)
            n = report(label, url, st)
            if not n:
                keys_of(st, f"{label} — что есть в состоянии", page.url)
            best = max(best, n)

        await browser.close()

    print("\n" + "-" * 60)
    if best:
        print(f"Подбор читается, карточек за запрос: до {best}.")
    else:
        print("Ни один адрес не отдал список вакансий — нужен другой путь.")
    return 0


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
