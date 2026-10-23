"""Разведчик телеграм-каналов: находит новые каналы с вакансиями и следит за старыми.

Зачем. Список каналов собирался руками и пополнялся, только когда человек
вспоминал об этом. Хорошие каналы появляются постоянно, а старые глохнут или
уходят в рекламу, и каждый лишний канал стоит денег на разборе каждого поста.

Раз в неделю (таймер jobsignal-tgscout):

1. ИЩЕТ кандидатов в трёх источниках:
   * упоминания в постах уже отслеживаемых каналов (t.me/… и @…): каналы
     с вакансиями постоянно ссылаются друг на друга, и это бесплатно —
     тексты уже в базе;
   * глобальный поиск Telegram по ключевым словам профилей (только чтение);
   * подписки человека: их он уже отобрал сам.
2. ПРОВЕРЯЕТ каждого по публичному превью t.me/s/<канал> (без аккаунта):
   жив ли, как часто пишет, и — главное — какая доля последних постов это
   вакансии ПОД НАШИ ПРОФИЛИ. «Публикует вакансии» и «публикует нужные
   вакансии» — разные вещи: первый ручной обход 07.10 нашёл 11 каналов
   с вакансиями, по профилю из них было 4. Решает модель (Haiku) по
   пятнадцати последним постам.
3. РЕШАЕТ: нужных вакансий в неделю не меньше TG_SCOUT_AUTO_MIN — канал
   добавляется сам (не больше TG_SCOUT_MAX_ADD за прогон); меньше, но не
   ноль — кандидат в дашборде (Каналы → кандидаты), решает человек; иначе
   отказ с причиной, повторная проверка через TG_SCOUT_RECHECK_DAYS дней.
4. СЛЕДИТ за отслеживаемыми: канал, добавленный разведчиком, который за
   30 дней не дал ни одной вакансии выше MATCH_THRESHOLD, отключается сам.
   Каналы, добавленные человеком, не трогаем — только называем в отчёте.
5. ОТЧЁТ в бот: что добавлено, что ждёт решения, что отключено.

Telegram-аккаунт используется только для чтения (поиск, список подписок)
и под общим замком сессии (tg_lock): сторож входящих в это время пропускает
прогон. Ничего не пишет, ни на что не подписывается.
"""
from __future__ import annotations

import asyncio
import html as htmlmod
import logging
import os
import re
import time
from collections import Counter
from datetime import datetime, timedelta, timezone

from bs4 import BeautifulSoup
from sqlalchemy import text as sql_text

from ..db import Channel, ChannelCandidate, get_session_factory, utcnow
from .collector import CollectorAgent, parse_channel_title, parse_messages

log = logging.getLogger("jobsignal")


def _env_f(name: str, default: float) -> float:
    try:
        return float(os.environ.get(name, default))
    except ValueError:
        return default


AUTO_MIN = _env_f("TG_SCOUT_AUTO_MIN", 1.5)       # нужных вакансий в неделю для автодобавления
PENDING_MIN = _env_f("TG_SCOUT_PENDING_MIN", 0.3)  # ниже — отказ
MIN_VAC_SHARE = _env_f("TG_SCOUT_MIN_VAC_SHARE", 0.4)  # доля вакансий среди постов
MAX_ADD = int(_env_f("TG_SCOUT_MAX_ADD", 5))
MAX_EVAL = int(_env_f("TG_SCOUT_MAX_EVAL", 40))
RECHECK_DAYS = int(_env_f("TG_SCOUT_RECHECK_DAYS", 90))
MENTION_DAYS = int(_env_f("TG_SCOUT_MENTION_DAYS", 60))
DISABLE_AFTER_DAYS = int(_env_f("TG_SCOUT_DISABLE_AFTER_DAYS", 30))
SAMPLE = 15
SAMPLE_CHARS = 500
FETCH_DELAY = 3.0

# Запросы к глобальному поиску Telegram. Переопределяются TG_SCOUT_QUERIES
# (через запятую).
DEFAULT_QUERIES = [
    "product owner вакансии", "продакт менеджер вакансии", "вакансии менеджер продукта",
    "fintech вакансии", "AI product manager", "IT вакансии senior",
    "вакансии CRM маркетинг", "программа лояльности вакансии", "retention вакансии",
    "маркетинг вакансии head",
]

SCOUT_SOURCE = "auto_scout"

# Имя канала/пользователя Telegram: 5-32 символа, буквы, цифры, подчёркивание.
_HANDLE = r"([A-Za-z][A-Za-z0-9_]{4,31})"
_TME_RE = re.compile(r"(?:https?://)?(?:t\.me|telegram\.me)/(?!s/|joinchat|c/|\+|addlist)"
                     + _HANDLE + r"(?![A-Za-z0-9_])")
_AT_RE = re.compile(r"(?<![A-Za-z0-9_.@/])@" + _HANDLE + r"(?![A-Za-z0-9_])")
# Не каналы: служебные адреса и боты.
_NOT_CHANNEL = {"share", "proxy", "socks", "iv", "setlanguage", "login", "joinchat"}


def _norm(h: str) -> str:
    return (h or "").strip().lstrip("@").lower()


# ── 1. кандидаты ─────────────────────────────────────────────────────────────

def mention_candidates(session, skip: set[str], limit: int) -> Counter:
    """Каналы, на которые ссылаются посты отслеживаемых каналов.

    Людей (рекрутёров) отсекаем по базе вакансий: их хэндлы мы уже знаем как
    контакты. Остальное отсеет проверка превью: у пользователя и у группы
    ленты постов на t.me/s/ нет.
    """
    since = datetime.now(timezone.utc) - timedelta(days=MENTION_DAYS)
    people = {_norm(h) for (h,) in session.execute(sql_text(
        "select distinct recruiter_handle from vacancies "
        "where recruiter_handle is not null")) if h}
    rows = session.execute(sql_text(
        "select r.text from raw_posts r join channels c on c.id=r.channel_id "
        "where coalesce(c.source,'') != 'hh' and r.text is not null "
        "and r.posted_at > :since"), {"since": since.replace(tzinfo=None)})
    cnt: Counter = Counter()
    for (text,) in rows:
        found = {_norm(m) for m in _TME_RE.findall(text)}
        found |= {_norm(m) for m in _AT_RE.findall(text)}
        for h in found:
            if (h in skip or h in people or h in _NOT_CHANNEL
                    or h.endswith("bot")):
                continue
            cnt[h] += 1
    # Канал, упомянутый один раз, чаще случайность (подпись автора, реклама).
    return Counter(dict((h, n) for h, n in cnt.most_common(limit) if n >= 2))


async def _telegram_candidates(queries: list[str], skip: set[str],
                               with_subs: bool) -> dict[str, dict]:
    from telethon import TelegramClient, functions

    api_id = int(os.environ.get("TG_API_ID") or 0)
    api_hash = (os.environ.get("TG_API_HASH") or "").strip()
    name = (os.environ.get("TG_INBOX_SESSION") or "data/tg_inbox").strip()
    out: dict[str, dict] = {}
    client = TelegramClient(name, api_id, api_hash)
    await client.start()
    try:
        for q in queries:
            try:
                res = await client(functions.contacts.SearchRequest(q=q, limit=30))
            except Exception as exc:  # noqa: BLE001 — один запрос не рушит поиск
                log.warning("[scout] поиск «%s» не выполнился: %s", q, exc)
                continue
            for ch in res.chats:
                h = _norm(getattr(ch, "username", "") or "")
                # Только каналы-ленты: у групп нет публичного превью постов.
                if h and getattr(ch, "broadcast", False) and h not in skip:
                    out.setdefault(h, {"title": getattr(ch, "title", "") or h,
                                       "subs": getattr(ch, "participants_count", 0) or 0,
                                       "source": "tg_search", "query": q})
            await asyncio.sleep(2)
        if with_subs:
            async for dialog in client.iter_dialogs():
                ent = dialog.entity
                h = _norm(getattr(ent, "username", "") or "")
                if (h and dialog.is_channel and getattr(ent, "broadcast", False)
                        and h not in skip):
                    out[h] = {"title": getattr(ent, "title", "") or h,
                              "subs": getattr(ent, "participants_count", 0) or 0,
                              "source": "my_subscriptions", "query": ""}
    finally:
        await client.disconnect()
    return out


def telegram_candidates(skip: set[str]) -> tuple[dict[str, dict], str]:
    """Кандидаты из поиска Telegram и подписок. Второе — причина, если не вышло."""
    name = (os.environ.get("TG_INBOX_SESSION") or "data/tg_inbox").strip()
    if not os.environ.get("TG_API_ID") or not os.path.exists(f"{name}.session"):
        return {}, "нет сессии Telegram (tools/tg_login.py) — поиск и подписки пропущены"
    raw = os.environ.get("TG_SCOUT_QUERIES", "")
    queries = [q.strip() for q in raw.split(",") if q.strip()] or DEFAULT_QUERIES
    from jobsignal.tg_lock import tg_session_lock
    with tg_session_lock(wait_s=180) as got:
        if not got:
            return {}, "сессия Telegram занята дольше 3 минут — поиск пропущен"
        try:
            return asyncio.run(_telegram_candidates(queries, skip, with_subs=True)), ""
        except Exception as exc:  # noqa: BLE001 — упоминания всё равно проверим
            log.exception("[scout] Telegram недоступен")
            return {}, f"Telegram недоступен: {type(exc).__name__}: {exc}"


# ── 2. проверка ──────────────────────────────────────────────────────────────

def _target_roles() -> str:
    """Целевые роли — из профилей (поле focus в config/profiles.yaml)."""
    custom = os.environ.get("TG_SCOUT_ROLES", "").strip()
    if custom:
        return custom
    try:
        from jobsignal import profiles
        text = profiles.focus_text()
    except Exception:  # noqa: BLE001
        text = ""
    return text or "продакт-менеджмент, уровень senior и выше"


JUDGE_SYSTEM = (
    "Ты оцениваешь телеграм-канал как источник вакансий для конкретного "
    "кандидата. Тебе дают название канала и пронумерованные последние посты.\n\n"
    "Для КАЖДОГО поста реши: vacancy — это объявление о вакансии (а не новость, "
    "реклама курса, статья, резюме соискателя, дайджест ссылок без описания); "
    "relevant — это вакансия, на которую кандидат реально может откликнуться "
    "по роли и уровню.\n\n"
    "Целевые роли кандидата: {roles}.\n\n"
    "Вакансии вне целевых ролей (разработка, дизайн, продажи, рядовые "
    "исполнители, джуны и стажёры) — не relevant, даже если в названии есть "
    "знакомое слово.\n\n"
    "Отвечай строго одним JSON без markdown:\n"
    '{{"posts": [{{"n": 1, "vacancy": true, "relevant": false}}], '
    '"niche": "<ниша канала 2-4 слова>", "summary": "<одна фраза: что за канал>"}}'
)


def _subscribers(html: str) -> int:
    soup = BeautifulSoup(html, "lxml")
    for el in soup.select(".tgme_channel_info_counter"):
        kind = el.select_one(".counter_type")
        val = el.select_one(".counter_value")
        if kind and val and "subscriber" in kind.get_text():
            raw = val.get_text(strip=True).upper().replace(" ", "")
            mult = 1000 if raw.endswith("K") else 1_000_000 if raw.endswith("M") else 1
            try:
                return int(float(raw.rstrip("KM")) * mult)
            except ValueError:
                return 0
    return 0


def evaluate(handle: str, html: str | None, judge=None) -> dict:
    """Метрики канала по превью. judge(system, user) → dict подменяется в тестах."""
    res = {"handle": handle, "ok": False}
    if not html:
        # Сеть, а не свойство канала: не записывать отказ на 90 дней.
        res["reason"] = "превью t.me/s/ не открылось"
        res["error"] = True
        return res
    posts = parse_messages(html)
    if not posts:
        res["reason"] = "нет публичной ленты (пользователь, группа или превью выключено)"
        return res
    res["title"] = parse_channel_title(html) or handle
    res["subs"] = _subscribers(html)
    posts.sort(key=lambda p: p["message_id"])
    dates = [p["posted_at"] for p in posts if p["posted_at"]]
    now = datetime.utcnow()
    if not dates:
        res["reason"] = "у постов нет дат"
        return res
    last_age = (now - max(dates)).total_seconds() / 86400
    span_days = max((max(dates) - min(dates)).total_seconds() / 86400, 1.0)
    per_week = len(dates) / span_days * 7
    res.update(last_age_days=round(last_age, 1), posts_per_week=round(per_week, 1),
               max_message_id=posts[-1]["message_id"])
    if last_age > 30:
        res["reason"] = f"молчит {last_age:.0f} дн."
        return res

    sample = posts[-SAMPLE:]
    user = f"Канал: {res['title']} (@{handle})\n\n" + "\n\n".join(
        f"[{i}] {p['text'][:SAMPLE_CHARS]}" for i, p in enumerate(sample, 1))
    system = JUDGE_SYSTEM.format(roles=_target_roles())
    if judge is None:
        from jobsignal.config import get_config
        from jobsignal.llm import complete_json
        model = get_config().settings.parser_model
        judge = lambda s, u: complete_json(s, u, model=model, max_tokens=900,  # noqa: E731
                                           tag="tg_scout", cache_system=True)
    try:
        data = judge(system, user)
    except Exception as exc:  # noqa: BLE001
        res["reason"] = f"модель не ответила: {exc}"
        res["error"] = True
        return res
    marks = {int(r.get("n", 0)): r for r in (data.get("posts") or [])
             if isinstance(r, dict)}
    vac = sum(1 for i in range(1, len(sample) + 1) if marks.get(i, {}).get("vacancy"))
    rel = sum(1 for i in range(1, len(sample) + 1)
              if marks.get(i, {}).get("vacancy") and marks.get(i, {}).get("relevant"))
    n = len(sample)
    res.update(ok=True, sample=n, vacancies=vac, relevant=rel,
               vac_share=round(vac / n, 2), rel_share=round(rel / n, 2),
               rel_per_week=round(per_week * rel / n, 2),
               niche=str(data.get("niche") or "")[:100],
               summary=str(data.get("summary") or "")[:300])
    return res


def decide(m: dict) -> tuple[str, str]:
    """(added | pending | rejected, причина по-человечески)."""
    if not m.get("ok"):
        return "rejected", m.get("reason", "не проверен")
    desc = (f"нужных вакансий ~{m['rel_per_week']:g}/нед "
            f"({m['relevant']} из {m['sample']} постов), вакансий {m['vac_share']:.0%}, "
            f"последний пост {m['last_age_days']:g} дн. назад")
    if (m["rel_per_week"] >= AUTO_MIN and m["vac_share"] >= MIN_VAC_SHARE
            and m["last_age_days"] <= 14):
        return "added", desc
    if m["rel_per_week"] >= PENDING_MIN:
        return "pending", desc
    return "rejected", desc


# ── 3. надзор за отслеживаемыми ──────────────────────────────────────────────

def audit(session, match_threshold: int) -> dict:
    """Что дают отслеживаемые каналы за 30 дней; отключить пустые авто-каналы."""
    now = datetime.now(timezone.utc).replace(tzinfo=None)
    since = now - timedelta(days=30)
    rows = session.execute(sql_text("""
        select c.id, c.handle, c.source, c.added_at,
          (select count(*) from raw_posts r where r.channel_id=c.id
             and r.posted_at > :since) posts,
          (select count(*) from vacancies v where v.channel_id=c.id
             and v.created_at > :since) vacs,
          (select count(*) from vacancies v where v.channel_id=c.id
             and v.created_at > :since
             and (select max(score) from match_scores m where m.vacancy_id=v.id) >= :thr) good
        from channels c
        where c.active = 1 and coalesce(c.source,'') != 'hh'
    """), {"since": since, "thr": match_threshold}).mappings().all()
    disabled, idle, silent = [], [], []
    for r in rows:
        added = r["added_at"]
        if isinstance(added, str):
            try:
                added = datetime.fromisoformat(added[:19].replace("T", " "))
            except ValueError:
                added = None
        old = added is None or (now - added).days >= DISABLE_AFTER_DAYS
        if not old:
            continue
        if r["source"] == SCOUT_SOURCE and r["good"] == 0:
            ch = session.get(Channel, r["id"])
            ch.active = False
            ch.enabled = False
            disabled.append((r["handle"], r["posts"], r["vacs"]))
        elif r["posts"] == 0:
            silent.append(r["handle"])
        elif r["good"] == 0:
            idle.append((r["handle"], r["posts"], r["vacs"]))
    session.commit()
    return {"checked": len(rows), "disabled": disabled, "idle": idle,
            "silent": silent}


# ── агент ────────────────────────────────────────────────────────────────────

class ChannelScoutAgent:
    name = "channel_scout"

    def __init__(self, dry_run: bool = False):
        self.dry_run = dry_run
        # fetch() конфиг не читает; None — чтобы не тянуть загрузку резюме.
        self._fetcher = CollectorAgent(None)

    def _skip_set(self, s) -> set[str]:
        skip = {_norm(h) for (h,) in s.query(Channel.handle).all() if h}
        cutoff = datetime.now(timezone.utc) - timedelta(days=RECHECK_DAYS)
        for c in s.query(ChannelCandidate).all():
            found = c.found_at
            if found is not None and found.tzinfo is None:
                found = found.replace(tzinfo=timezone.utc)
            # Ожидающих и добавленных не трогаем; отказ — пересмотр по сроку.
            if c.status != "rejected" or (found and found > cutoff):
                skip.add(_norm(c.handle))
        return skip

    def run(self) -> dict:
        from jobsignal.config import get_config
        s = get_session_factory()()
        try:
            skip = self._skip_set(s)
            mentions = mention_candidates(s, skip, limit=MAX_EVAL)
            tg, tg_problem = telegram_candidates(skip)

            cands: dict[str, dict] = {}
            for h, n in mentions.items():
                cands[h] = {"source": "mentions", "query": f"упоминаний: {n}",
                            "weight": n}
            for h, info in tg.items():
                prev = cands.get(h)
                cands[h] = {**info, "weight": (prev or {}).get("weight", 0) + 1}
                if prev:
                    cands[h]["query"] = f"{info.get('query', '')}; {prev['query']}"
            # Подписки и частые упоминания — вперёд: они вероятнее годны.
            order = sorted(cands, key=lambda h: (cands[h]["source"] != "my_subscriptions",
                                                 -cands[h]["weight"]))[:MAX_EVAL]
            log.info("[scout] кандидатов: упоминания %d, Telegram %d; проверяю %d",
                     len(mentions), len(tg), len(order))

            results = {"added": [], "pending": [], "rejected": [], "errors": 0}
            for h in order:
                html = self._fetcher.fetch(h)
                m = evaluate(h, html)
                verdict, why = decide(m)
                if m.get("error"):
                    results["errors"] += 1
                    continue          # модель не ответила — проверим в следующий раз
                if verdict == "added" and len(results["added"]) >= MAX_ADD:
                    verdict = "pending"
                    why += f" (лимит {MAX_ADD} автодобавлений за прогон)"
                results[verdict].append((h, m, why))
                log.info("[scout] @%s → %s: %s", h, verdict, why)
                if not self.dry_run:
                    self._store(s, h, cands[h], m, verdict, why)
                time.sleep(FETCH_DELAY)

            thr = int(get_config().settings.match_threshold)
            aud = audit(s, thr) if not self.dry_run else {
                "checked": 0, "disabled": [], "idle": [], "silent": []}
        finally:
            s.close()

        report = {"agent": self.name, "dry_run": self.dry_run,
                  "candidates": len(cands), "evaluated": len(order),
                  "added": [h for h, _, _ in results["added"]],
                  "pending": [h for h, _, _ in results["pending"]],
                  "rejected": len(results["rejected"]),
                  "errors": results["errors"], "telegram_problem": tg_problem,
                  "audit": aud}
        if not self.dry_run:
            self._notify(results, aud, tg_problem)
        return report

    def _store(self, s, h: str, cand: dict, m: dict, verdict: str, why: str) -> None:
        title = m.get("title") or cand.get("title") or h
        subs = m.get("subs") or cand.get("subs") or 0
        desc = why + (f". {m['summary']}" if m.get("summary") else "")
        row = s.query(ChannelCandidate).filter_by(handle=h).first()
        if row is None:
            row = ChannelCandidate(handle=h)
            s.add(row)
        row.title = title[:200]
        row.description = desc[:1000]
        row.subscriber_count = subs
        row.source = f"scout:{cand.get('source', '')}"
        row.search_query = (cand.get("query") or "")[:200]
        row.found_at = utcnow()
        row.status = verdict
        if verdict == "added" and not s.query(Channel).filter_by(handle=h).first():
            # last_message_id не ставим: первый сбор возьмёт ленту превью
            # (около 20 последних постов), и вакансии канала появятся сразу.
            s.add(Channel(handle=h, title=title[:200], niche=m.get("niche") or "",
                          active=True, enabled=True, source=SCOUT_SOURCE,
                          subscriber_count=subs, verified_at=utcnow(),
                          post_count_30d=int(m.get("posts_per_week", 0) * 30 / 7)))
        s.commit()

    def _notify(self, results: dict, aud: dict, tg_problem: str) -> None:
        if not (results["added"] or results["pending"] or aud["disabled"]
                or aud["idle"] or aud["silent"] or tg_problem):
            return
        esc = htmlmod.escape
        lines = ["🔎 <b>Разведка каналов</b>"]
        if results["added"]:
            lines.append(f"\n<b>Добавлены в отслеживание ({len(results['added'])}):</b>")
            for h, m, why in results["added"]:
                lines.append(f"• @{esc(h)} — {esc(m.get('niche') or '')}\n  {esc(why)}")
        if results["pending"]:
            lines.append(f"\n<b>Ждут решения ({len(results['pending'])})</b> — "
                         "в кабинете: Каналы → кандидаты:")
            for h, m, why in results["pending"][:10]:
                lines.append(f"• @{esc(h)}: {esc(why)}")
        if aud["disabled"]:
            lines.append(f"\n<b>Отключены</b> (добавлены разведчиком, за "
                         f"{DISABLE_AFTER_DAYS} дн. ни одной подходящей вакансии):")
            for h, posts, vacs in aud["disabled"]:
                lines.append(f"• @{esc(h)}: постов {posts}, вакансий {vacs}")
        if aud["idle"]:
            lines.append("\nДобавлены вручную, но за 30 дн. без подходящих вакансий: "
                         + ", ".join(f"@{esc(h)}" for h, _, _ in aud["idle"][:15])
                         + " — можно отключить в кабинете")
        if aud["silent"]:
            lines.append("Молчат 30 дн.: " + ", ".join(f"@{esc(h)}" for h in aud["silent"][:15]))
        if tg_problem:
            lines.append(f"\n⚠️ {esc(tg_problem)}")
        try:
            from .notify_bot import _send
            _send("\n".join(lines)[:4000])
        except Exception as exc:  # noqa: BLE001
            log.warning("[scout] отчёт в бот не отправлен: %s", exc)
