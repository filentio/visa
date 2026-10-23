"""Разведка: как на hh выбрать нужное резюме при отклике. Ничего не отправляет.

Повод. 23.10 у кандидата на hh два резюме — Product Owner и лояльность/CRM.
Автоотклик должен в форме отклика выбирать резюме того профиля, под который
вакансия набрала лучший балл; иначе hh отправит резюме по умолчанию. Как
устроен выбор на живой странице, отсюда не видно, а гадать с отправкой
откликов нельзя. Поэтому сначала разведка:

  1. /applicant/resumes — какие резюме есть, их названия и хэши, и какое
     резюме подходит к какому профилю по hh_resume_title из profiles.yaml;
  2. страница вакансии — куда ведёт кнопка «Откликнуться» (только читаем
     атрибут, НЕ нажимаем) и что про резюме лежит в состоянии страницы
     (JSON HH-Lux-InitialState);
  3. только с флагом --open-form: открыть адрес формы отклика (GET, без
     нажатий) и выписать, как на ней выбирается резюме. Делать это на
     вакансии, на которую вы и так готовы откликнуться.

Вывод и снимки — в data/probe_resume_choice/. Пришлите вывод целиком.

    cd /opt/jobsignal_local && .venv/bin/python tools/probe_hh_resume_choice.py
    cd /opt/jobsignal_local && .venv/bin/python tools/probe_hh_resume_choice.py \\
        https://hh.ru/vacancy/123456 --open-form
"""
from __future__ import annotations

import asyncio
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from dotenv import load_dotenv  # noqa: E402
load_dotenv("config/.env")

OUT = Path("data/probe_resume_choice")
STATE_RE = re.compile(r'id="HH-Lux-InitialState"[^>]*>(.*?)</template>', re.S)
RESUME_HREF_RE = re.compile(r"/resume/([0-9a-f]{20,})")


def _walk(obj, path="", out=None, depth=0):
    """Пути в JSON, где ключ или строковое значение говорит о резюме."""
    if out is None:
        out = []
    if depth > 12 or len(out) > 80:
        return out
    if isinstance(obj, dict):
        for k, v in obj.items():
            p = f"{path}.{k}" if path else k
            if "resume" in k.lower():
                out.append((p, json.dumps(v, ensure_ascii=False)[:300]))
            _walk(v, p, out, depth + 1)
    elif isinstance(obj, list):
        for i, v in enumerate(obj[:20]):
            _walk(v, f"{path}[{i}]", out, depth + 1)
    return out


def _pick_vacancy() -> str | None:
    """Вакансия hh из очереди — с лучшим баллом под действующие профили."""
    import sqlite3
    from jobsignal import profiles
    names = profiles.names()
    if not names:
        return None
    c = sqlite3.connect("data/jobsignal.db")
    marks = ",".join("?" * len(names))
    row = c.execute(
        f"select v.link from vacancies v join match_scores m on m.vacancy_id=v.id "
        f"where v.contact_type='hh' and v.link like '%hh.ru/vacancy/%' "
        f"and m.profile_key in ({marks}) order by m.score desc, v.id desc limit 1",
        names).fetchone()
    return row[0] if row else None


async def main() -> int:
    from playwright.async_api import async_playwright
    from jobsignal import hh_session, profiles

    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    open_form = "--open-form" in sys.argv
    vacancy = args[0] if args else _pick_vacancy()
    OUT.mkdir(parents=True, exist_ok=True)
    path = hh_session.state_path()
    print(f"сессия: {path} ({'есть' if path.exists() else 'НЕТ'})")
    print("профили:")
    for p in profiles.all_profiles():
        print(f"  {p.key}: «{p.name}», ищем на hh резюме со словами «{p.hh_resume_title or '—'}»")
    print()

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=True,
                                           proxy=hh_session.playwright_proxy())
        ctx = await browser.new_context(
            storage_state=str(path) if path.exists() else None,
            locale="ru-RU", viewport={"width": 1440, "height": 900})
        ctx.set_default_timeout(60000)
        await hh_session.block_heavy(ctx, "probe")
        page = await ctx.new_page()

        # 1. Резюме аккаунта
        print("=== 1. резюме на hh (/applicant/resumes)")
        await page.goto("https://hh.ru/applicant/resumes", wait_until="domcontentloaded")
        await asyncio.sleep(8)
        await page.screenshot(path=str(OUT / "1-resumes.png"), full_page=True)
        found = {}
        for a in await page.locator("a[href*='/resume/']").all():
            href = await a.get_attribute("href") or ""
            m = RESUME_HREF_RE.search(href)
            text = (await a.inner_text()).strip().replace("\n", " ")
            if m and text and len(text) > 3:
                found.setdefault(m.group(1), text[:120])
        if not found:
            print("  резюме не найдены — проверь сессию (run.py hh-session)")
        for h, title in found.items():
            prof = [p.key for p in profiles.all_profiles()
                    if p.hh_resume_title and p.hh_resume_title.lower() in title.lower()]
            print(f"  {h}  «{title}»  → профиль: {', '.join(prof) or 'НЕ СОПОСТАВЛЕНО'}")
        for p in profiles.all_profiles():
            hits = [h for h, t in found.items()
                    if p.hh_resume_title and p.hh_resume_title.lower() in t.lower()]
            if len(hits) != 1:
                print(f"  ВНИМАНИЕ: профилю {p.key} подходит резюме: {len(hits)} "
                      f"(нужно ровно одно) — поправь hh_resume_title в profiles.yaml")
        print()

        if not vacancy:
            print("нет вакансии для шагов 2-3: укажи адрес вакансии hh аргументом")
            await browser.close()
            return 0

        # 2. Страница вакансии — только чтение
        print(f"=== 2. вакансия {vacancy} (ничего не нажимаем)")
        await page.goto(vacancy, wait_until="domcontentloaded")
        await asyncio.sleep(8)
        await page.screenshot(path=str(OUT / "2-vacancy.png"), full_page=False)
        hrefs = []
        for sel in ('[data-qa="vacancy-response-link-top"]',
                    '[data-qa="vacancy-response-link-bottom"]'):
            loc = page.locator(sel)
            if await loc.count():
                el = loc.first
                tag = await el.evaluate("e => e.tagName")
                href = await el.get_attribute("href")
                print(f"  {sel}: <{tag}> href={href}")
                if href:
                    hrefs.append(href)
        html = await page.content()
        (OUT / "2-vacancy.html").write_text(html, encoding="utf-8")
        m = STATE_RE.search(html)
        if m:
            try:
                state = json.loads(m.group(1))
                paths = _walk(state)
                print(f"  в состоянии страницы про резюме: {len(paths)} мест")
                for p_, v in paths[:40]:
                    print(f"    {p_} = {v}")
            except ValueError as exc:
                print(f"  состояние страницы не разобрано: {exc}")
        else:
            print("  блока HH-Lux-InitialState нет")
        print()

        # 3. Форма отклика — только по явному флагу, только GET
        if open_form and hrefs:
            url = hrefs[0] if hrefs[0].startswith("http") else "https://hh.ru" + hrefs[0]
            print(f"=== 3. форма отклика {url} (GET, без нажатий)")
            await page.goto(url, wait_until="domcontentloaded")
            await asyncio.sleep(8)
            await page.screenshot(path=str(OUT / "3-form.png"), full_page=True)
            (OUT / "3-form.html").write_text(await page.content(), encoding="utf-8")
            print(f"  адрес после загрузки: {page.url}")
            qa = await page.evaluate("""() => [...document.querySelectorAll('[data-qa]')]
                .map(e => e.getAttribute('data-qa'))
                .filter(q => /resume|response|submit|letter/i.test(q))""")
            print(f"  data-qa про резюме/отклик: {sorted(set(qa))[:60]}")
            radios = await page.evaluate("""() => [...document.querySelectorAll(
                'input[type=radio], input[type=checkbox], select, option')]
                .map(e => ({tag: e.tagName, name: e.name || '', value: (e.value || '').slice(0, 60),
                            checked: !!e.checked,
                            label: ((e.closest('label') || e.parentElement || {}).innerText || '')
                                   .trim().slice(0, 100)}))""")
            for r in radios[:30]:
                print(f"    {r}")
            print()
        elif open_form:
            print("=== 3. у кнопки отклика нет адреса — форма открывается скриптом; "
                  "пришли вывод, разберёмся по снимку и состоянию страницы")

        await browser.close()
    print(f"снимки и HTML: {OUT}/")
    return 0


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
