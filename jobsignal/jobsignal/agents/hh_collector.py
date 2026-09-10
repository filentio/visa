"""
HHCollector — сбор вакансий с hh.ru через встроенное состояние страницы.

api.hh.ru/vacancies с этого сервера закрыт: ddos-guard отдаёт 403 на любой
запрос к разделу вакансий — проверено семью способами, список в
docs/ARCHITECTURE.md. Повторно проверять его не нужно.

Страницы самого hh.ru при этом открываются обычным requests и без сессии, а
внутри несут то же структурированное состояние, что раздавал API:
<template id="HH-Lux-InitialState"> с JSON. Берём данные оттуда, а не из
вёрстки — переименование классов и перестановка карточек нам безразличны.

Результаты складываются в raw_posts с channel_id специального HH-канала,
дальше идёт тот же пайплайн: parser → dedup → matcher.

Молчаливых нулей здесь быть не должно. Коллектор возвращал «добавлено 0» с
конца июня, конвейер шёл дальше, и поломку не замечали два с половиной
месяца. Поэтому пропавший блок состояния, изменившийся путь к ключам,
неудача всех запросов подряд или ноль карточек по всем поискам — это
HHStructureChanged, а не пустой результат.
"""
from __future__ import annotations

import html as htmlmod
import json
import logging
import os
import re
import time
from datetime import datetime, timezone
from typing import Optional

import requests

from jobsignal.db import get_session_factory, Channel, RawPost

log = logging.getLogger("jobsignal")

SEARCH_URL = "https://hh.ru/search/vacancy"
HH_VACANCY_URL = "https://hh.ru/vacancy/{}"

# Блок состояния в HTML. Он же — единственное, что может «поехать» при
# редизайне: имя шаблона и путь к ключам ниже.
STATE_RE = re.compile(r'id="HH-Lux-InitialState"[^>]*>(.*?)</template>', re.S)

HEADERS = {
    # Сайт отдаёт страницы и роботу, но с UA инструмента здороваться незачем:
    # это обычный просмотр каталога, а не обращение к API.
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ru-RU,ru;q=0.9",
}

# Сколько страниц выдачи забирать на каждый поиск (50 вакансий на странице).
HH_PAGES = int(os.environ.get("HH_PAGES", "1"))
# Потолок догрузок описаний за прогон: это по HTTP-запросу на вакансию.
HH_DESC_LIMIT = int(os.environ.get("HH_DESC_LIMIT", "150"))
# Паузы, чтобы не выглядеть перебором каталога.
PAGE_DELAY = float(os.environ.get("HH_PAGE_DELAY", "1.5"))
DESC_DELAY = float(os.environ.get("HH_DESC_DELAY", "0.7"))

# Поисковые запросы под три профиля
DEFAULT_SEARCHES = [
    # CPO / Head of Product
    {"text": "CPO", "area": 1, "experience": "between3And6"},
    {"text": "Chief Product Officer", "area": 1},
    {"text": "Head of Product", "area": 1, "experience": "between3And6"},
    {"text": "Директор по продукту", "area": 1},
    # Senior AI PM
    {"text": "AI Product Manager", "area": 1},
    {"text": "ML Product Manager", "area": 1},
    {"text": "Product Manager AI", "area": 1},
    # Senior PM fintech
    {"text": "Product Manager fintech", "area": 1},
    {"text": "Продакт менеджер финтех", "area": 1},
    {"text": "Product Owner банк", "area": 1},
    # General senior PM
    {"text": "Senior Product Manager", "area": 1, "experience": "between3And6"},
    {"text": "Lead Product Manager", "area": 1},
]

# area=1 = Москва, area=2 = Санкт-Петербург, area=113 = Россия
# professional_roles: 96=Продуктовый менеджер, 104=Руководитель
PM_PROFESSIONAL_ROLES = [96, 104, 157]  # PM, IT-директор, Бизнес-аналитик

# Коды опыта у hh — в состоянии лежит код, человеку нужен текст.
EXPERIENCE_NAMES = {
    "noExperience": "Без опыта",
    "between1And3": "1-3 года",
    "between3And6": "3-6 лет",
    "moreThan6": "Более 6 лет",
}
WORK_FORMAT_NAMES = {
    "REMOTE": "удалённо",
    "ON_SITE": "в офисе",
    "HYBRID": "гибрид",
    "FIELD_WORK": "разъездная",
}


class HHStructureChanged(RuntimeError):
    """hh отдал страницу, но ожидаемых данных в ней нет.

    Отдельный тип, чтобы вызывающий отличал поломку разбора от честного
    «сегодня по этому запросу ничего не нашлось».
    """


def _hh_channel(session) -> Channel:
    """Получить или создать виртуальный канал для hh.ru вакансий."""
    ch = session.query(Channel).filter_by(handle="hh_ru").first()
    if not ch:
        ch = Channel(
            handle="hh_ru",
            title="hh.ru (API)",
            niche="hh",
            active=True,
            source="hh",
        )
        session.add(ch)
        session.commit()
        session.refresh(ch)
    return ch


def _extract_state(html: str, where: str) -> dict:
    """JSON-состояние страницы. Нет блока или он не парсится — это поломка."""
    m = STATE_RE.search(html)
    if not m:
        raise HHStructureChanged(
            f"{where}: блок HH-Lux-InitialState не найден "
            f"(страница {len(html)} симв.) — hh сменил разметку состояния"
        )
    try:
        return json.loads(htmlmod.unescape(m.group(1).strip()))
    except json.JSONDecodeError as exc:
        raise HHStructureChanged(
            f"{where}: состояние найдено, но не разбирается как JSON: {exc}"
        ) from exc


def _clean_html(text: Optional[str]) -> Optional[str]:
    if not text:
        return None
    plain = re.sub(r"<[^>]+>", " ", text)
    plain = htmlmod.unescape(re.sub(r"\s+", " ", plain)).strip()
    return plain or None


def _salary(compensation: dict | None) -> Optional[dict]:
    """compensation страницы → тот же вид, что отдавал API."""
    if not compensation or compensation.get("noCompensation"):
        return None
    frm, to = compensation.get("from"), compensation.get("to")
    if frm is None and to is None:
        return None
    return {"from": frm, "to": to,
            "currency": compensation.get("currencyCode") or "RUR"}


def _card_to_item(card: dict) -> Optional[dict]:
    """Карточка выдачи → dict в форме, которую уже умеет _format_post_text."""
    vac_id = card.get("vacancyId")
    if not vac_id:
        return None
    company = card.get("company") or {}
    area = card.get("area") or {}
    formats = []
    for block in card.get("workFormats") or []:
        for code in block.get("workFormatsElement") or []:
            formats.append(WORK_FORMAT_NAMES.get(code, code.lower()))
    published = card.get("creationTime")
    if not published:
        pub = card.get("publicationTime") or {}
        published = pub.get("$") if isinstance(pub, dict) else None
    return {
        "id": vac_id,
        "name": card.get("name"),
        "employer": {"name": company.get("name") or company.get("visibleName")},
        "salary": _salary(card.get("compensation")),
        "area": {"name": area.get("name")},
        "schedule": {"name": ", ".join(formats)} if formats else {},
        "experience": {"name": EXPERIENCE_NAMES.get(card.get("workExperience"), "")},
        "snippet": {},
        "published_at": published,
    }


def _fetch_search_page(params: dict, page: int) -> list[dict]:
    """Одна страница выдачи. Пустой список — законный ответ, а не сбой."""
    p = {**params, "page": page, "items_on_page": 50}
    r = requests.get(SEARCH_URL, params=p, headers=HEADERS, timeout=25)
    r.raise_for_status()

    state = _extract_state(r.text, f"поиск '{params.get('text')}' стр. {page}")
    result = state.get("vacancySearchResult")
    if not isinstance(result, dict):
        raise HHStructureChanged(
            f"поиск '{params.get('text')}': в состоянии нет vacancySearchResult "
            f"(есть ключи: {sorted(state)[:8]}…) — путь к данным изменился"
        )
    cards = result.get("vacancies")
    if not isinstance(cards, list):
        raise HHStructureChanged(
            f"поиск '{params.get('text')}': vacancySearchResult.vacancies не "
            f"список ({type(cards).__name__}) — путь к данным изменился"
        )
    return cards


def _fetch_description(vac_id: int) -> Optional[str]:
    """Полный текст вакансии со страницы. Без него матчер судит по названию."""
    r = requests.get(HH_VACANCY_URL.format(vac_id), headers=HEADERS, timeout=25)
    r.raise_for_status()
    state = _extract_state(r.text, f"вакансия {vac_id}")
    view = state.get("vacancyView")
    if not isinstance(view, dict):
        raise HHStructureChanged(
            f"вакансия {vac_id}: в состоянии нет vacancyView — "
            f"путь к данным изменился"
        )
    return _clean_html(view.get("description"))


def _with_description(text: str, desc: str) -> str:
    """Дописать описание в уже собранный текст поста — перед строкой ссылки."""
    lines = text.split("\n")
    idx = next((i for i, l in enumerate(lines) if l.startswith("Ссылка:")),
               len(lines))
    lines.insert(idx, f"Описание: {desc}")
    return "\n".join(lines)


def _gone(exc: Exception) -> bool:
    """Вакансия снята: описания не будет и ждать его больше незачем."""
    resp = getattr(exc, "response", None)
    return resp is not None and resp.status_code in (404, 410)


def _format_post_text(item: dict) -> str:
    """Convert hh.ru vacancy item to text for parser."""
    parts = []

    name = item.get("name", "")
    if name:
        parts.append(name)

    employer = item.get("employer", {})
    if employer.get("name"):
        parts.append(f"Компания: {employer['name']}")

    salary = item.get("salary")
    if salary:
        frm = salary.get("from")
        to = salary.get("to")
        cur = salary.get("currency", "RUB")
        if frm and to:
            parts.append(f"Зарплата: {frm}–{to} {cur}")
        elif frm:
            parts.append(f"Зарплата: от {frm} {cur}")
        elif to:
            parts.append(f"Зарплата: до {to} {cur}")

    area = item.get("area", {})
    if area.get("name"):
        parts.append(f"Локация: {area['name']}")

    schedule = item.get("schedule", {})
    if schedule.get("name"):
        parts.append(f"График: {schedule['name']}")

    experience = item.get("experience", {})
    if experience.get("name"):
        parts.append(f"Опыт: {experience['name']}")

    snippet = item.get("snippet", {})
    if snippet.get("requirement"):
        req = snippet["requirement"].replace("<highlighttext>", "").replace("</highlighttext>", "")
        parts.append(f"Требования: {req}")
    if snippet.get("responsibility"):
        resp = snippet["responsibility"].replace("<highlighttext>", "").replace("</highlighttext>", "")
        parts.append(f"Обязанности: {resp}")

    # Полное описание со страницы вакансии: без него матчер оценивал hh-вакансии
    # по одному названию должности.
    if item.get("description"):
        parts.append(f"Описание: {item['description']}")

    url = HH_VACANCY_URL.format(item.get("id", ""))
    parts.append(f"Ссылка: {url}")

    return "\n".join(parts)


class HHCollector:
    def __init__(self, searches: Optional[list[dict]] = None):
        self._sf = get_session_factory()
        self.searches = searches or DEFAULT_SEARCHES

    def run(self) -> dict:
        session = self._sf()
        ch = _hh_channel(session)

        # existing hh vacancy IDs to avoid duplicates
        existing = {
            row[0]
            for row in session.query(RawPost.tg_message_id)
            .filter(RawPost.channel_id == ch.id)
            .all()
        }

        added = 0
        seen_ids: set[int] = set()
        fresh: list[dict] = []
        cards_total = 0
        requests_failed = 0

        for search_params in self.searches:
            found_here = 0
            for page in range(HH_PAGES):
                try:
                    cards = _fetch_search_page(search_params, page)
                except HHStructureChanged:
                    session.close()
                    raise           # поломка разбора — наверх, без «добавлено 0»
                except requests.RequestException as exc:
                    requests_failed += 1
                    log.warning("[hh] '%s' стр. %d: запрос не прошёл: %s",
                                search_params.get("text"), page, exc)
                    break

                cards_total += len(cards)
                found_here += len(cards)
                for card in cards:
                    item = _card_to_item(card)
                    if not item:
                        continue
                    vac_id = int(item["id"])
                    if vac_id in seen_ids or vac_id in existing:
                        continue
                    seen_ids.add(vac_id)
                    fresh.append(item)

                if not cards:
                    break           # дальше страниц нет
                time.sleep(PAGE_DELAY)

            log.info("[hh] '%s' → %d карточек", search_params.get("text"), found_here)

        # Ноль по одному запросу — бывает. Ноль по всем сразу или полный отказ
        # сети означает поломку: именно так выглядел мёртвый сбор с конца июня.
        if requests_failed == len(self.searches):
            session.close()
            raise HHStructureChanged(
                f"ни один из {len(self.searches)} запросов к hh.ru не прошёл — "
                f"сбор не работает"
            )
        if cards_total == 0:
            session.close()
            raise HHStructureChanged(
                f"все {len(self.searches)} поисков вернули пустую выдачу — "
                f"это поломка сбора, а не отсутствие вакансий"
            )

        log.info("[hh] карточек: %d, новых: %d — догружаю описания",
                 cards_total, len(fresh))
        descriptions = self._settle_descriptions(session, ch, fresh)

        for item in fresh:
            posted_at = self._posted_at(item)
            post = RawPost(
                channel_id=ch.id,
                tg_message_id=int(item["id"]),  # hh vacancy id как уникальный id
                text=_format_post_text(item),
                post_url=HH_VACANCY_URL.format(item["id"]),
                posted_at=posted_at,
                parsed=False,
                # Описание не влезло в бюджет прогона — пост сохраняем, но
                # парсеру не отдаём: оценка по одному заголовку бессмысленна.
                # Следующий прогон доберёт описание и снимет признак.
                awaiting_details=not item.get("description"),
            )
            session.add(post)
            added += 1

        try:
            session.commit()
        except Exception as exc:
            session.rollback()
            session.close()
            raise                      # молча терять собранное нельзя

        session.close()
        result = {"agent": "hh_collector", "added": added,
                  "searches": len(self.searches), "cards": cards_total,
                  "descriptions": descriptions,
                  "requests_failed": requests_failed}
        log.info("[hh] добавлено вакансий: %d (описаний загружено: %d)",
                 added, descriptions)
        return result

    def _settle_descriptions(self, session, ch, fresh: list[dict]) -> int:
        """Раздать бюджет догрузок: сначала долги прошлых прогонов, потом новые.

        Порядок именно такой. Иначе свежая выдача каждый прогон вытесняла бы
        накопленные долги, и вакансии, не влезшие в бюджет, висели бы без
        описания вечно — а без описания матчер оценивает их по заголовку.
        """
        budget = HH_DESC_LIMIT
        loaded = failed = 0

        pending = (
            session.query(RawPost)
            .filter(RawPost.channel_id == ch.id,
                    RawPost.awaiting_details == True)  # noqa: E712
            .order_by(RawPost.id)
            .limit(budget)
            .all()
        )
        if pending:
            log.info("[hh] долгов по описаниям: %d", len(pending))
        for post in pending:
            if budget <= 0:
                break
            budget -= 1
            try:
                desc = _fetch_description(int(post.tg_message_id))
            except HHStructureChanged:
                session.close()
                raise
            except requests.RequestException as exc:
                if _gone(exc):
                    # Вакансия снята: описания не будет, держать её в долгах
                    # незачем — пусть парсер разбирает то, что есть.
                    post.awaiting_details = False
                else:
                    failed += 1
                    log.warning("[hh] описание %s не загрузилось: %s",
                                post.tg_message_id, exc)
                time.sleep(DESC_DELAY)
                continue
            if desc:
                post.text = _with_description(post.text, desc)
                loaded += 1
            # Пустое описание — тоже ответ: ждать больше нечего.
            post.awaiting_details = False
            time.sleep(DESC_DELAY)
        session.commit()

        for item in fresh:
            if budget <= 0:
                break
            budget -= 1
            try:
                item["description"] = _fetch_description(int(item["id"]))
                loaded += 1
            except HHStructureChanged:
                session.close()
                raise
            except requests.RequestException as exc:
                failed += 1
                log.warning("[hh] описание %s не загрузилось: %s",
                            item["id"], exc)
            time.sleep(DESC_DELAY)

        if fresh and loaded == 0 and failed >= 5:
            session.close()
            raise HHStructureChanged(
                f"ни одно описание не загрузилось ({failed} попыток) — "
                f"страницы вакансий недоступны"
            )
        waiting = sum(1 for i in fresh if not i.get("description"))
        if waiting:
            log.info("[hh] без описания пока остаются %d новых — доберём "
                     "следующими прогонами, в матчер они не пойдут", waiting)
        return loaded

    @staticmethod
    def _posted_at(item: dict) -> datetime:
        published = item.get("published_at")
        if not published:
            return datetime.now(timezone.utc)
        try:
            return datetime.fromisoformat(published.replace("Z", "+00:00"))
        except Exception:
            return datetime.now(timezone.utc)
