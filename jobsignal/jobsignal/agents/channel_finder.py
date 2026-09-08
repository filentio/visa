"""
ChannelFinder — Модуль B:
  1. Ищет новые каналы через tgstat.ru и telemetr.me
  2. Проверяет канал (публичный? живой? есть посты за 30 дней?)
  3. Кладёт кандидатов в channel_candidates (статус pending)
  4. По одобрению из дашборда добавляет в channels
"""
from __future__ import annotations

import logging
import re
import time
from datetime import datetime, timezone, timedelta
from typing import Optional

import requests
from bs4 import BeautifulSoup

from jobsignal.db import get_session_factory, Channel, ChannelCandidate

log = logging.getLogger("jobsignal")

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
}

# Keywords we search for on tgstat / telemetr
DEFAULT_QUERIES = [
    "product manager вакансии",
    "CPO вакансии",
    "AI product manager",
    "fintech вакансии продукт",
    "IT вакансии продакт",
    "head of product вакансии",
    "product owner jobs",
]


# ── tgstat search ─────────────────────────────────────────────────────────────

def _search_tgstat(query: str, max_results: int = 10) -> list[dict]:
    """Search tgstat.ru for channels matching query."""
    url = "https://tgstat.ru/search"
    params = {"q": query, "type": "channel"}
    candidates = []
    try:
        r = requests.get(url, params=params, headers=HEADERS, timeout=15)
        if r.status_code != 200:
            log.warning("[finder] tgstat %s → HTTP %s", query, r.status_code)
            return []
        soup = BeautifulSoup(r.text, "lxml")
        # tgstat renders channel cards with class "peer-item"
        for card in soup.select(".peer-item")[:max_results]:
            handle_el = card.select_one(".peer-item-username")
            title_el = card.select_one(".peer-item-title")
            subs_el = card.select_one(".peer-item-members")
            desc_el = card.select_one(".peer-item-description")
            if not handle_el:
                continue
            handle = handle_el.get_text(strip=True).lstrip("@")
            if not handle:
                continue
            subs = _parse_number(subs_el.get_text(strip=True) if subs_el else "0")
            candidates.append({
                "handle": handle,
                "title": title_el.get_text(strip=True) if title_el else handle,
                "description": desc_el.get_text(strip=True) if desc_el else "",
                "subscriber_count": subs,
                "source": "tgstat",
                "search_query": query,
            })
    except Exception as exc:
        log.warning("[finder] tgstat error for %r: %s", query, exc)
    return candidates


# ── telemetr search ───────────────────────────────────────────────────────────

def _search_telemetr(query: str, max_results: int = 10) -> list[dict]:
    """Search telemetr.me for channels matching query."""
    url = "https://telemetr.me/en/channels"
    params = {"q": query}
    candidates = []
    try:
        r = requests.get(url, params=params, headers=HEADERS, timeout=15)
        if r.status_code != 200:
            log.warning("[finder] telemetr %s → HTTP %s", query, r.status_code)
            return []
        soup = BeautifulSoup(r.text, "lxml")
        for card in soup.select(".channel-card, .channel-item")[:max_results]:
            link_el = card.select_one("a[href*='t.me'], a[href*='/channel/']")
            title_el = card.select_one(".channel-name, h3, h4")
            subs_el = card.select_one(".subscribers, .members")
            desc_el = card.select_one(".channel-description, .description")
            if not link_el:
                continue
            href = link_el.get("href", "")
            handle = re.search(r"t\.me/([A-Za-z0-9_]+)", href)
            if not handle:
                handle = re.search(r"/channel/([A-Za-z0-9_]+)", href)
            if not handle:
                continue
            handle = handle.group(1)
            subs = _parse_number(subs_el.get_text(strip=True) if subs_el else "0")
            candidates.append({
                "handle": handle,
                "title": title_el.get_text(strip=True) if title_el else handle,
                "description": desc_el.get_text(strip=True) if desc_el else "",
                "subscriber_count": subs,
                "source": "telemetr",
                "search_query": query,
            })
    except Exception as exc:
        log.warning("[finder] telemetr error for %r: %s", query, exc)
    return candidates


# ── channel health check ──────────────────────────────────────────────────────

def verify_channel(handle: str) -> dict:
    """
    Check channel via t.me/s/{handle}.
    Returns: {alive, post_count_30d, latest_post_date, title}
    """
    handle = handle.lstrip("@")
    url = f"https://t.me/s/{handle}"
    result = {
        "alive": False,
        "post_count_30d": 0,
        "latest_post_date": None,
        "title": handle,
    }
    try:
        r = requests.get(url, headers=HEADERS, timeout=15)
        if r.status_code != 200:
            return result
        soup = BeautifulSoup(r.text, "lxml")
        # title
        title_el = soup.select_one(".tgme_channel_info_header_title")
        if title_el:
            result["title"] = title_el.get_text(strip=True)
        # posts
        posts = soup.select(".tgme_widget_message")
        if not posts:
            return result
        result["alive"] = True
        # count posts in last 30 days
        cutoff = datetime.now(timezone.utc) - timedelta(days=30)
        recent = 0
        latest = None
        for p in posts:
            time_el = p.select_one("time[datetime]")
            if time_el:
                try:
                    dt_str = time_el["datetime"]
                    dt = datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
                    if latest is None or dt > latest:
                        latest = dt
                    if dt >= cutoff:
                        recent += 1
                except Exception:
                    pass
        result["post_count_30d"] = recent
        result["latest_post_date"] = latest
    except Exception as exc:
        log.warning("[finder] verify %s error: %s", handle, exc)
    return result


# ── main runner ───────────────────────────────────────────────────────────────

class ChannelFinder:
    def __init__(self, queries: Optional[list[str]] = None):
        self._sf = get_session_factory()
        self.queries = queries or DEFAULT_QUERIES

    def run(self, queries: Optional[list[str]] = None) -> dict:
        queries = queries or self.queries
        session = self._sf()
        new_candidates = 0
        skipped_existing = 0

        # existing handles (channels + candidates) to avoid duplicates
        existing_handles = {
            row[0].lstrip("@").lower()
            for row in session.query(Channel.handle).all()
        }
        existing_candidates = {
            row[0].lstrip("@").lower()
            for row in session.query(ChannelCandidate.handle).all()
        }
        known = existing_handles | existing_candidates

        for query in queries:
            log.info("[finder] searching: %r", query)
            found: list[dict] = []
            found += _search_tgstat(query, max_results=8)
            time.sleep(2)
            found += _search_telemetr(query, max_results=8)
            time.sleep(2)

            for c in found:
                h = c["handle"].lstrip("@").lower()
                if h in known:
                    skipped_existing += 1
                    continue

                # health check
                log.info("[finder] verifying @%s …", c["handle"])
                health = verify_channel(c["handle"])
                time.sleep(1)

                if not health["alive"]:
                    log.info("[finder] @%s not alive, skip", c["handle"])
                    continue
                if health["post_count_30d"] < 2:
                    log.info(
                        "[finder] @%s only %d posts/30d, skip",
                        c["handle"], health["post_count_30d"],
                    )
                    continue

                candidate = ChannelCandidate(
                    handle=c["handle"],
                    title=health.get("title") or c["title"],
                    description=c.get("description", ""),
                    subscriber_count=c.get("subscriber_count", 0),
                    source=c["source"],
                    search_query=c["search_query"],
                )
                session.add(candidate)
                known.add(h)
                new_candidates += 1
                log.info(
                    "[finder] candidate @%s (%d posts/30d)",
                    c["handle"], health["post_count_30d"],
                )

            try:
                session.commit()
            except Exception as exc:
                session.rollback()
                log.error("[finder] db error: %s", exc)

        session.close()
        result = {
            "agent": "channel_finder",
            "new_candidates": new_candidates,
            "skipped_existing": skipped_existing,
        }
        log.info("[finder] done: %s", result)
        return result

    def add_channel_from_candidate(self, candidate_id: int, niche: str = "product") -> bool:
        """Promote a candidate to active channel (called from dashboard)."""
        session = self._sf()
        try:
            cand = session.get(ChannelCandidate, candidate_id)
            if not cand:
                return False
            existing = session.query(Channel).filter_by(
                handle=cand.handle
            ).first()
            if existing:
                cand.status = "added"
                session.commit()
                return True
            ch = Channel(
                handle=cand.handle,
                title=cand.title,
                niche=niche,
                active=True,
                source=cand.source,
                subscriber_count=cand.subscriber_count,
                verified_at=datetime.now(timezone.utc),
            )
            session.add(ch)
            cand.status = "added"
            session.commit()
            log.info("[finder] added channel @%s", cand.handle)
            return True
        except Exception as exc:
            session.rollback()
            log.error("[finder] add_channel error: %s", exc)
            return False
        finally:
            session.close()

    def reject_candidate(self, candidate_id: int) -> bool:
        session = self._sf()
        try:
            cand = session.get(ChannelCandidate, candidate_id)
            if not cand:
                return False
            cand.status = "rejected"
            session.commit()
            return True
        finally:
            session.close()

    def add_manual(
        self, handle: str, title: str = "", niche: str = "product"
    ) -> dict:
        """Add channel manually. Verifies it first."""
        handle = handle.lstrip("@")
        session = self._sf()
        try:
            existing = session.query(Channel).filter_by(handle=handle).first()
            if existing:
                return {"ok": False, "error": f"@{handle} уже в списке"}

            health = verify_channel(handle)
            if not health["alive"]:
                return {"ok": False, "error": f"@{handle} не найден или закрыт"}

            ch = Channel(
                handle=handle,
                title=title or health.get("title", handle),
                niche=niche,
                active=True,
                source="manual",
                post_count_30d=health["post_count_30d"],
                verified_at=datetime.now(timezone.utc),
            )
            session.add(ch)
            session.commit()
            log.info("[finder] manual add @%s", handle)
            return {
                "ok": True,
                "handle": handle,
                "title": ch.title,
                "post_count_30d": health["post_count_30d"],
            }
        except Exception as exc:
            session.rollback()
            return {"ok": False, "error": str(exc)}
        finally:
            session.close()


# ── helpers ───────────────────────────────────────────────────────────────────

def _parse_number(text: str) -> int:
    """Parse '12.3K', '1 234', '1M' → int."""
    text = text.strip().replace(" ", "").replace("\xa0", "")
    m = re.search(r"([\d.,]+)\s*([KkМмMm]?)", text)
    if not m:
        return 0
    num = float(m.group(1).replace(",", "."))
    suffix = m.group(2).upper()
    if suffix in ("K", "К"):
        num *= 1_000
    elif suffix in ("M", "М"):
        num *= 1_000_000
    return int(num)
