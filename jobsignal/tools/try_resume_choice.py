"""Проверить выбор резюме в живой форме отклика hh. Отклик НЕ отправляется.

Что делает: открывает вакансию, по состоянию страницы определяет резюме
профиля (как боевой прогон), нажимает «Откликнуться», выбирает резюме тем же
кодом, что и автоотклик, печатает, какое резюме показано в форме до и после,
сохраняет снимок и разметку формы и закрывает форму крестиком. Кнопку
отправки не нажимает никогда.

Защита от случайного отклика: hh отправляет отклик прямо на нажатии
«Откликнуться», если отправить можно только одно резюме и вопросов нет.
Поэтому нажимаем, только когда hh предлагает для вакансии два резюме и
больше — тогда он всегда показывает форму.

    cd /opt/jobsignal_local && .venv/bin/python tools/try_resume_choice.py \\
        https://hh.ru/vacancy/138965445 [--profile po]
"""
from __future__ import annotations

import argparse
import asyncio
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from dotenv import load_dotenv  # noqa: E402
load_dotenv("config/.env")

OUT = Path("data/try_resume_choice")


async def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("url")
    ap.add_argument("--profile", default=None,
                    help="ключ профиля (po, loyalty); по умолчанию — лучший балл вакансии")
    opts = ap.parse_args()

    from jobsignal import hh_resumes, profiles
    from jobsignal.agents.hh_apply import SEL, HHApplyAgent
    from jobsignal.config import get_config

    m = re.search(r"/vacancy/(\d+)", opts.url)
    if not m:
        print("нужен адрес вида https://hh.ru/vacancy/123")
        return 2
    key = m.group(1)
    prof = profiles.by_key(opts.profile) if opts.profile else None
    if prof is None:
        import sqlite3
        c = sqlite3.connect("data/jobsignal.db")
        row = c.execute(
            "select m.profile_key from vacancies v join match_scores m on m.vacancy_id=v.id "
            "where v.link like ? order by m.score desc limit 1", (f"%{key}%",)).fetchone()
        prof = profiles.resolve(row[0] if row else None)
    print(f"профиль: {prof.key} «{prof.name}»")
    OUT.mkdir(parents=True, exist_ok=True)

    agent = HHApplyAgent(get_config(), dry_run=True)
    agent.screenshots_dir = OUT
    item = {"id": 0, "key": key, "url": f"https://hh.ru/vacancy/{key}",
            "role": "", "company": "", "score": 0, "profile": prof.name}
    async with agent._browser() as page:
        agent._page = page
        if not await agent._logged_in(page):
            print(f"сессия не работает: {agent._session_detail}")
            return 1
        await agent._load_resume_map(page)
        print("резюме по профилям:", {k: v["title"] for k, v in agent._resume_map.items()})

        for attempt in (1, 2, 3):
            try:
                await page.goto(item["url"], wait_until="domcontentloaded")
                break
            except Exception as exc:  # noqa: BLE001 — прокси медленный
                print(f"вакансия не открылась (попытка {attempt}/3): "
                      f"{str(exc).splitlines()[0]}")
                if attempt == 3:
                    print("прокси не отвечает — проверь остаток трафика у провайдера")
                    return 1
        await page.wait_for_timeout(3000)
        if await agent._visible(page, SEL["already_applied"], timeout=1.5):
            print("на эту вакансию отклик уже есть — возьми другую")
            return 0
        vr = hh_resumes.vacancy_resumes(hh_resumes.page_state(await page.content()), key)
        if vr is None:
            print("в состоянии страницы нет списка резюме — нажимать не буду")
            return 1
        print(f"hh предлагает резюме: {vr.unused}, уже использованы: {vr.used}")
        if len(vr.unused) < 2:
            print("доступно меньше двух резюме — hh может отправить отклик прямо "
                  "на нажатии, поэтому не нажимаю. Возьми другую вакансию.")
            return 0
        choice, refusal = await agent._pick_resume(page, item)
        if refusal is not None:
            print(f"выбирать нечего: {refusal.result.value} — {refusal.detail}")
            return 0
        print(f"нужно резюме: «{choice['title']}» (id {choice['id']}); "
              f"узнаём по: {choice['marks']}, чужие: {choice['others']}")

        button = await agent._visible(page, SEL["respond_button"], timeout=5)
        if button is None:
            print("кнопки «Откликнуться» нет")
            return 1
        await button.click()
        await page.wait_for_timeout(2500)
        if await agent._visible(page, SEL["already_applied"], timeout=2):
            print("!!! hh отправил отклик прямо на нажатии — такого при двух резюме "
                  "не ожидалось. Проверь отклик на hh.")
            return 1
        reloc = await agent._visible(page, SEL["relocation_confirm"], timeout=2)
        if reloc is not None:
            await reloc.click()
            await page.wait_for_timeout(800)
        print(f"адрес после нажатия: {page.url}")
        print(f"в форме до выбора: {await agent._shown_resumes(page)}")
        await page.screenshot(path=str(OUT / "1-form-before.png"), full_page=True)
        (OUT / "1-form-before.html").write_text(await page.content(), encoding="utf-8")

        ok = await agent._select_resume(page, choice)
        print(f"выбор подтверждён: {'ДА' if ok else 'НЕТ'}")
        print(f"в форме после выбора: {await agent._shown_resumes(page)}")
        await page.screenshot(path=str(OUT / "2-form-after.png"), full_page=True)
        (OUT / "2-form-after.html").write_text(await page.content(), encoding="utf-8")

        close = await agent._visible(page, '[data-qa="response-popup-close"]', timeout=2)
        if close is not None:
            await close.click()
            print("форма закрыта крестиком, отклик не отправлен")
        else:
            print("крестика нет (отдельная страница отклика) — просто ухожу, "
                  "отклик не отправлен")
    print(f"снимки и разметка: {OUT}/")
    return 0


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
