"""BoardsCollector — вакансии с площадок помимо hh: Хабр Карьера и getmatch.

Зачем. До 08.11 источников было два: hh и телеграм-каналы. Разведка
(tools/probe_sources.py) показала, что серверу без прокси отвечают JSON-ом:

  * career.habr.com/api/frontend/vacancies — список с поиском; полный текст
    вакансии — в ld+json JobPosting на её странице;
  * getmatch.ru/api/offers — список с описанием прямо в карточке; запрос
    q площадка, похоже, не учитывает (на «product owner» — все 928).

Поэтому отбор по названию делаем сами, ДО разбора моделью: в названии должно
быть слово из профилей (matcher.KEEP_RE: продукт, owner, CRM, лояльность...).
Иначе каждая чужая вакансия стоила бы вызова модели на разборе.

Как с hh: карточка → raw_posts в виртуальном канале площадки → parser →
matcher → уведомление в телеграм со ссылкой. Откликаются на площадке руками
(автоотклик — только hh). Тип контакта вакансии — habr / getmatch
(outreach.BOARD_PATTERNS по ссылке).

Сбой одной площадки не роняет другие и конвейер: ошибка в журнал, счётчик в
итог. Площадка, у которой за прогон не прошёл ни один запрос, — уровнем error.

Настройки (config/.env):
  BOARDS                 какие площадки собирать, по умолчанию habr,getmatch
  BOARDS_EVERY_H         не чаще чем раз в N часов (по умолчанию 3)
  BOARDS_MAX_AGE_DAYS    вакансии не старше (по умолчанию 21)
  BOARDS_PAGES_PER_RUN   страниц вакансий Хабра за прогон (по умолчанию 60)
"""
from __future__ import annotations

import html as htmlmod
import json
import logging
import os
import re
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import quote

import requests

from jobsignal.db import Channel, RawPost, get_session_factory

log = logging.getLogger("jobsignal")

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36")
HEADERS = {"User-Agent": UA, "Accept-Language": "ru-RU,ru;q=0.9",
           "Accept": "application/json, text/html;q=0.9"}

BOARDS = [b.strip() for b in os.environ.get("BOARDS", "habr,getmatch").split(",") if b.strip()]
EVERY_H = float(os.environ.get("BOARDS_EVERY_H", "3"))
MAX_AGE_DAYS = int(os.environ.get("BOARDS_MAX_AGE_DAYS", "21"))
PAGES_PER_RUN = int(os.environ.get("BOARDS_PAGES_PER_RUN", "60"))
STATE_FILE = Path(os.environ.get("BOARDS_STATE_FILE", "data/boards_state.json"))
DELAY = 1.0
DESC_CAP = 6000

_TAG_RE = re.compile(r"<[^>]+>")
_LD_RE = re.compile(r'<script[^>]+application/ld\+json[^>]*>(.*?)</script>', re.S)


def _text(html_or_text: str | None) -> str:
    t = _TAG_RE.sub(" ", html_or_text or "")
    return " ".join(htmlmod.unescape(t).split())


def _relevant(title: str) -> bool:
    from jobsignal.agents.matcher import KEEP_RE
    return bool(title and KEEP_RE.search(title))


def _queries() -> list[str]:
    """Запросы — из профилей (те же, что для hh), без дублей."""
    from jobsignal import profiles
    out, seen = [], set()
    for p in profiles.all_profiles():
        for q in p.searches:
            if q.lower() not in seen:
                seen.add(q.lower())
                out.append(q)
    return out or ["Product Owner", "Менеджер продукта", "Руководитель CRM"]


def _fresh(dt: str | None) -> bool:
    if not dt:
        return True
    try:
        d = datetime.fromisoformat(dt.replace("Z", "+00:00"))
    except ValueError:
        return True
    if d.tzinfo is None:
        d = d.replace(tzinfo=timezone.utc)
    return d >= datetime.now(timezone.utc) - timedelta(days=MAX_AGE_DAYS)


def _get(url: str):
    r = requests.get(url, headers=HEADERS, timeout=25)
    r.raise_for_status()
    return r


# ── Хабр Карьера ─────────────────────────────────────────────────────────────

def _habr_list(q: str) -> list[dict]:
    out = []
    for page in (1, 2):
        d = _get("https://career.habr.com/api/frontend/vacancies"
                 f"?q={quote(q)}&type=all&sort=date&per_page=25&page={page}").json()
        items = d.get("list") or []
        out += items
        meta = d.get("meta") or {}
        if not items or page >= int(meta.get("totalPages") or 1):
            break
        time.sleep(DELAY)
    return out


def _habr_description(href: str) -> str:
    html = _get("https://career.habr.com" + href).text
    for m in _LD_RE.finditer(html):
        try:
            data = json.loads(m.group(1))
        except ValueError:
            continue
        for d in (data if isinstance(data, list) else [data]):
            if isinstance(d, dict) and d.get("@type") == "JobPosting":
                return _text(d.get("description"))
    return ""


def _habr_item(it: dict, desc: str) -> dict:
    comp = (it.get("company") or {}).get("title") or ""
    sal = (it.get("salary") or {}).get("formatted") or ""
    locs = ", ".join(x.get("title", "") for x in it.get("locations") or [])
    if it.get("remoteWork"):
        locs = (locs + ", " if locs else "") + "можно удалённо"
    skills = ", ".join(x.get("title", "") for x in it.get("skills") or [])
    return {"id": int(it["id"]), "title": it.get("title") or "", "company": comp,
            "salary": sal, "location": locs, "skills": skills,
            "level": it.get("qualification") or "",
            "date": (it.get("publishedDate") or {}).get("date"),
            "url": "https://career.habr.com" + (it.get("href") or f"/vacancies/{it['id']}"),
            "description": desc}


# ── getmatch ────────────────────────────────────────────────────────────────

def _getmatch_list(q: str) -> list[dict]:
    d = _get(f"https://getmatch.ru/api/offers?q={quote(q)}&limit=100").json()
    return d.get("offers") or []


def _getmatch_item(it: dict) -> dict:
    comp = ((it.get("company") or {}).get("name") if isinstance(it.get("company"), dict)
            else it.get("company_name")) or ""
    sal = ""
    if not it.get("salary_hidden"):
        lo, hi = it.get("salary_display_from"), it.get("salary_display_to")
        cur = it.get("salary_currency") or ""
        if lo and hi:
            sal = f"{lo}–{hi} {cur}".strip()
        elif lo or hi:
            sal = (f"от {lo}" if lo else f"до {hi}") + f" {cur}".rstrip()
    locs = []
    for loc in it.get("location_requirements") or []:
        city, fmt = loc.get("city") or "", loc.get("format") or ""
        locs.append(f"{city} ({fmt})" if fmt else city)
    skills = ", ".join(s.get("title") or s.get("name") or ""
                       for s in it.get("skills_objects") or [] if isinstance(s, dict))
    url = it.get("url") or f"/vacancies/{it.get('id')}"
    return {"id": int(it["id"]), "title": it.get("position") or "", "company": comp,
            "salary": sal, "location": "; ".join(locs[:4]), "skills": skills,
            "level": "", "date": it.get("published_at"),
            "url": url if url.startswith("http") else "https://getmatch.ru" + url,
            "description": _text(it.get("offer_description") or it.get("description_html"))}


# ── общее ────────────────────────────────────────────────────────────────────

SOURCE_TITLES = {"habr": "Хабр Карьера", "getmatch": "getmatch"}


def _post_text(item: dict, source: str) -> str:
    parts = [f"Источник: {SOURCE_TITLES.get(source, source)}",
             f"Вакансия: {item['title']}"]
    for label, key in (("Компания", "company"), ("Зарплата", "salary"),
                       ("Локация", "location"), ("Уровень", "level"), ("Навыки", "skills")):
        if item.get(key):
            parts.append(f"{label}: {item[key]}")
    if item.get("description"):
        parts.append(f"Описание: {item['description'][:DESC_CAP]}")
    parts.append(f"Ссылка: {item['url']}")
    return "\n".join(parts)


def _channel(session, source: str) -> Channel:
    handle = f"board_{source}"
    ch = session.query(Channel).filter_by(handle=handle).first()
    if not ch:
        ch = Channel(handle=handle, title=SOURCE_TITLES.get(source, source),
                     niche=source, active=True, source="board")
        session.add(ch)
        session.commit()
        session.refresh(ch)
    return ch


def _due(source: str) -> bool:
    try:
        last = json.loads(STATE_FILE.read_text()).get(source, 0)
    except (OSError, ValueError):
        last = 0
    return time.time() - float(last) >= EVERY_H * 3600 - 300


def _mark(source: str) -> None:
    try:
        state = json.loads(STATE_FILE.read_text())
    except (OSError, ValueError):
        state = {}
    state[source] = time.time()
    STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
    STATE_FILE.write_text(json.dumps(state))


class BoardsCollector:
    def __init__(self, boards: list[str] | None = None, force: bool = False):
        self.boards = boards or BOARDS
        self.force = force
        self._sf = get_session_factory()

    def run(self) -> dict:
        report = {}
        for source in self.boards:
            if source not in SOURCE_TITLES:
                log.warning("[boards] неизвестная площадка %s — пропуск", source)
                continue
            if not self.force and not _due(source):
                report[source] = "не время"
                continue
            try:
                report[source] = self._run_source(source)
                _mark(source)
            except Exception as exc:  # noqa: BLE001 — площадка не роняет конвейер
                log.error("[boards] %s: сбор сломан: %s", source, exc, exc_info=True)
                report[source] = f"ошибка: {exc}"
        log.info("[boards] итог: %s", report)
        return report

    def _run_source(self, source: str) -> dict:
        session = self._sf()
        try:
            ch = _channel(session, source)
            existing = {r[0] for r in session.query(RawPost.tg_message_id)
                        .filter(RawPost.channel_id == ch.id).all()}
            seen: set[int] = set()
            new_items: list[dict] = []
            ok = failed = cards = irrelevant = old = 0
            for q in _queries():
                try:
                    raw = _habr_list(q) if source == "habr" else _getmatch_list(q)
                    ok += 1
                except requests.RequestException as exc:
                    failed += 1
                    log.warning("[boards] %s '%s': запрос не прошёл: %s", source, q, exc)
                    continue
                cards += len(raw)
                for it in raw:
                    try:
                        vid = int(it["id"])
                    except (KeyError, TypeError, ValueError):
                        continue
                    if vid in seen or vid in existing:
                        continue
                    seen.add(vid)
                    title = it.get("title") or it.get("position") or ""
                    if it.get("archived") or it.get("is_active") is False:
                        continue
                    date = ((it.get("publishedDate") or {}).get("date")
                            if source == "habr" else it.get("published_at"))
                    if not _fresh(date):
                        old += 1
                        continue
                    if not _relevant(title):
                        irrelevant += 1
                        continue
                    new_items.append(it)
                time.sleep(DELAY)
                # getmatch, похоже, не фильтрует по q: один запрос даёт всю
                # ленту, остальные — те же карточки. Хватит одного.
                if source == "getmatch":
                    break

            if ok == 0 and failed:
                raise RuntimeError(f"ни один из {failed} запросов не прошёл")

            added = pages = 0
            for it in new_items:
                if source == "habr":
                    if pages >= PAGES_PER_RUN:
                        break
                    try:
                        desc = _habr_description(it.get("href") or f"/vacancies/{it['id']}")
                    except requests.RequestException as exc:
                        log.warning("[boards] habr #%s: страница не открылась: %s",
                                    it.get("id"), exc)
                        desc = ""
                    pages += 1
                    item = _habr_item(it, desc)
                    time.sleep(DELAY)
                else:
                    item = _getmatch_item(it)
                posted = None
                if item.get("date"):
                    try:
                        posted = datetime.fromisoformat(item["date"].replace("Z", "+00:00"))
                    except ValueError:
                        posted = None
                session.add(RawPost(channel_id=ch.id, tg_message_id=item["id"],
                                    text=_post_text(item, source), post_url=item["url"],
                                    posted_at=posted))
                added += 1
            session.commit()
            res = {"запросов": ok, "карточек": cards, "новых подходящих": len(new_items),
                   "добавлено": added, "не по профилю": irrelevant, "старые": old}
            log.info("[boards] %s: %s", source, res)
            return res
        finally:
            session.close()
