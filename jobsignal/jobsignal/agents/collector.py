"""Collector (Этап 1).

Читает публичное веб-превью каналов t.me/s/{handle} — без api_id, без аккаунта,
без файла сессии. Парсит сообщения, инкрементально складывает новые в raw_posts
(по tg_message_id), уважает rate-лимит t.me (паузы + обработка 429).

Работает только для ПУБЛИЧНЫХ каналов с включённым превью. Приватные/закрытые
каналы здесь не читаются — для них потребовался бы MTProto (Telethon + api_id).
"""
from __future__ import annotations

import logging
import time
from datetime import datetime, timezone

import requests
from bs4 import BeautifulSoup
from sqlalchemy import select

from ..db import Channel, RawPost, get_session_factory, utcnow
from .base import BaseAgent

log = logging.getLogger("jobsignal")

TME_URL = "https://t.me/s/{handle}"
USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"
)
REQUEST_DELAY = 3.0   # пауза между каналами, сек (вежливо к rate-лимиту по IP)
TIMEOUT = 20
MAX_RETRIES = 3


# --- Парсинг HTML веб-превью (вынесено отдельно, чтобы можно было тестировать) ---

def extract_text(el) -> str:
    """Текст сообщения с сохранением переносов строк (job-посты структурированы)."""
    for br in el.find_all("br"):
        br.replace_with("\n")
    text = el.get_text()
    # схлопываем хвостовые пробелы по строкам
    lines = [ln.rstrip() for ln in text.splitlines()]
    return "\n".join(lines).strip()


def parse_datetime(value: str) -> datetime | None:
    try:
        dt = datetime.fromisoformat(value)
        if dt.tzinfo is not None:
            dt = dt.astimezone(timezone.utc).replace(tzinfo=None)
        return dt
    except (ValueError, TypeError):
        return None


def parse_messages(html: str) -> list[dict]:
    """Возвращает [{message_id, text, posted_at, post_url, links}] из t.me/s/."""
    soup = BeautifulSoup(html, "lxml")
    out: list[dict] = []
    for div in soup.select("div.tgme_widget_message[data-post]"):
        data_post = div.get("data-post", "")
        parts = data_post.rsplit("/", 1)
        if len(parts) != 2 or not parts[1].isdigit():
            continue
        msg_id = int(parts[1])
        post_url = f"https://t.me/{data_post}"

        text_el = div.select_one("div.tgme_widget_message_text")
        text = extract_text(text_el) if text_el else ""
        if not text.strip():
            continue  # медиа без подписи / служебные — пропускаем

        # ссылки из тела поста (href): hh.ru, формы, t.me/username и т.п.
        links = []
        if text_el:
            for a in text_el.find_all("a", href=True):
                href = a["href"]
                if href.startswith("http") and href not in links:
                    links.append(href)
        if links:
            text += "\n\nСсылки в посте: " + " | ".join(links)

        posted_at = None
        time_el = div.select_one("a.tgme_widget_message_date time[datetime]")
        if time_el and time_el.get("datetime"):
            posted_at = parse_datetime(time_el["datetime"])

        out.append({"message_id": msg_id, "text": text, "posted_at": posted_at,
                    "post_url": post_url, "links": links})
    return out


def parse_channel_title(html: str) -> str | None:
    soup = BeautifulSoup(html, "lxml")
    el = soup.select_one(".tgme_channel_info_header_title, .tgme_header_title")
    return el.get_text(strip=True) if el else None


# --- Агент ---

class CollectorAgent(BaseAgent):
    name = "collector"

    def run(self) -> dict:
        Session = get_session_factory()
        total_new = 0
        scanned = 0
        with Session() as s:
            channels = (
                s.execute(select(Channel).where(Channel.active.is_(True)))
                .scalars()
                .all()
            )
            if not channels:
                log.warning("[collector] нет включённых каналов — заполни channels.yaml "
                            "и выполни seed-channels")
            for ch in channels:
                try:
                    new = self._collect_channel(s, ch)
                    total_new += new
                    scanned += 1
                    log.info("[collector] %s: +%d новых", ch.handle, new)
                except Exception as exc:  # noqa: BLE001
                    log.exception("[collector] %s: ошибка %s", ch.handle, exc)
                time.sleep(REQUEST_DELAY)
            s.commit()
        log.info("[collector] просканировано каналов: %d, всего новых постов: %d",
                 scanned, total_new)
        return {"agent": self.name, "channels": scanned, "new_posts": total_new}

    def _collect_channel(self, session, ch: Channel) -> int:
        handle = ch.handle.lstrip("@")
        html = self.fetch(handle)
        if html is None:
            return 0

        posts = parse_messages(html)
        if not posts:
            log.warning("[collector] %s: постов не найдено (приватный канал или "
                        "превью отключено)", ch.handle)
            ch.last_scanned_at = utcnow()
            return 0

        if not ch.title:
            title = parse_channel_title(html)
            if title:
                ch.title = title

        new_count = 0
        max_id = ch.last_message_id or 0
        for p in sorted(posts, key=lambda x: x["message_id"]):
            if ch.last_message_id and p["message_id"] <= ch.last_message_id:
                continue
            exists = session.scalar(
                select(RawPost).where(
                    RawPost.channel_id == ch.id,
                    RawPost.tg_message_id == p["message_id"],
                )
            )
            if exists:
                continue
            session.add(
                RawPost(
                    channel_id=ch.id,
                    tg_message_id=p["message_id"],
                    text=p["text"],
                    post_url=p.get("post_url"),
                    posted_at=p["posted_at"],
                    parsed=False,
                )
            )
            new_count += 1
            max_id = max(max_id, p["message_id"])

        ch.last_message_id = max_id
        ch.last_scanned_at = utcnow()
        return new_count

    def fetch(self, handle: str) -> str | None:
        """GET t.me/s/{handle} с ретраями и обработкой 429. Переопределяемо в тестах."""
        url = TME_URL.format(handle=handle)
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                resp = requests.get(
                    url,
                    headers={"User-Agent": USER_AGENT, "Accept-Language": "ru,en;q=0.9"},
                    timeout=TIMEOUT,
                )
                if resp.status_code == 200:
                    return resp.text
                if resp.status_code == 429:
                    wait = int(resp.headers.get("Retry-After", 30))
                    log.warning("[collector] %s: 429 rate-limit, ждём %dс", handle, wait)
                    time.sleep(wait)
                    continue
                log.warning("[collector] %s: HTTP %d", handle, resp.status_code)
                return None
            except requests.RequestException as exc:
                backoff = 2 ** attempt
                log.warning("[collector] %s: попытка %d/%d, сеть: %s (ждём %dс)",
                            handle, attempt, MAX_RETRIES, exc, backoff)
                time.sleep(backoff)
        log.error("[collector] %s: не удалось скачать после %d попыток", handle, MAX_RETRIES)
        return None
