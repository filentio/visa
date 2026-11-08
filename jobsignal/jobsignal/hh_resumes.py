"""Резюме аккаунта hh.ru и выбор нужного при отклике.

Повод. С 23.10 у кандидата на hh два резюме под два профиля (Product Owner и
лояльность/CRM). Отклик должен уходить с резюме того профиля, под который
вакансия набрала лучший балл; иначе hh подставит резюме по умолчанию.

Что известно по разведке 03.11 (tools/probe_hh_resume_choice.py). В состоянии
страницы вакансии (JSON HH-Lux-InitialState) hh сам пишет, какие резюме можно
отправить на эту вакансию:

    applicantVacancyResponseStatuses.<id вакансии>:
        resumes          {id резюме: {_attributes: {hash, ...}, ...}}
        usedResumeIds    уже откликались этими резюме
        unusedResumeIds  можно откликнуться
        hiddenResumeIds, unfinishedResumeIds

Поэтому решение принимается ДО клика по «Откликнуться»:
  * нужным резюме уже откликались — отклик есть;
  * нужного нет среди доступных — не отправляем, причина в отчёт;
  * доступно только нужное — hh отправит именно его, выбирать нечего;
  * доступно несколько — в форме отклика выбираем нужное и без
    подтверждённого выбора не отправляем.
После отправки страница перечитывается, и в usedResumeIds должно появиться
именно нужное резюме. Если ушло другое, боевые отклики встают на паузу
(файл BLOCK_FILE) до ручной проверки: run.py hh-resumes --unblock.

Какое резюме к какому профилю — по hh_resume_title из config/profiles.yaml:
часть названия резюме на hh, которая должна совпасть ровно с одним резюме.
Названия берём со страницы /applicant/resumes (кэш на сутки, чтобы не тратить
трафик прокси на каждый прогон).
"""
from __future__ import annotations

import json
import logging
import os
import re
import time
from dataclasses import dataclass, field
from pathlib import Path

log = logging.getLogger("jobsignal")

STATE_RE = re.compile(r'id="HH-Lux-InitialState"[^>]*>(.*?)</template>', re.S)
RESUME_HREF_RE = re.compile(r"/resume/([0-9a-f]{20,})")
CACHE_FILE = Path(os.environ.get("HH_RESUMES_CACHE", "data/hh_resumes.json"))
CACHE_HOURS = float(os.environ.get("HH_RESUMES_CACHE_HOURS", "24"))
BLOCK_FILE = Path(os.environ.get("HH_RESUME_BLOCK_FILE", "data/hh_resume_block.json"))


# ── состояние страницы ───────────────────────────────────────────────────────

def page_state(html: str) -> dict | None:
    m = STATE_RE.search(html or "")
    if not m:
        return None
    try:
        return json.loads(m.group(1))
    except ValueError:
        return None


def _title(obj) -> str:
    """Название резюме из объекта состояния: hh хранит его то строкой, то
    списком {"string": ...}."""
    t = obj.get("title") if isinstance(obj, dict) else None
    if isinstance(t, str):
        return t.strip()
    if isinstance(t, dict):
        return str(t.get("string") or t.get("value") or "").strip()
    if isinstance(t, list) and t:
        first = t[0]
        if isinstance(first, dict):
            return str(first.get("string") or first.get("value") or "").strip()
        return str(first).strip()
    return ""


@dataclass
class VacancyResumes:
    """Что hh разрешает отправить на конкретную вакансию."""
    resumes: dict[str, dict] = field(default_factory=dict)   # id → {hash, title}
    used: list[str] = field(default_factory=list)
    unused: list[str] = field(default_factory=list)
    hidden: list[str] = field(default_factory=list)
    unfinished: list[str] = field(default_factory=list)

    def id_by_hash(self, h: str) -> str | None:
        for rid, r in self.resumes.items():
            if r.get("hash") == h:
                return rid
        return None


def vacancy_resumes(state: dict | None, vacancy_key: str) -> VacancyResumes | None:
    if not state:
        return None
    statuses = state.get("applicantVacancyResponseStatuses") or {}
    st = statuses.get(str(vacancy_key))
    if not isinstance(st, dict):
        return None
    out = VacancyResumes(
        used=[str(x) for x in st.get("usedResumeIds") or []],
        unused=[str(x) for x in st.get("unusedResumeIds") or []],
        hidden=[str(x) for x in st.get("hiddenResumeIds") or []],
        unfinished=[str(x) for x in st.get("unfinishedResumeIds") or []],
    )
    for rid, r in (st.get("resumes") or {}).items():
        attrs = (r or {}).get("_attributes") or {}
        out.resumes[str(rid)] = {"hash": attrs.get("hash") or "",
                                 "title": _title(r)}
    return out


# ── резюме аккаунта и сопоставление с профилями ─────────────────────────────

def read_cache() -> dict | None:
    try:
        data = json.loads(CACHE_FILE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    if time.time() - float(data.get("at") or 0) > CACHE_HOURS * 3600:
        return None
    return data.get("resumes") or None


def drop_cache() -> None:
    try:
        CACHE_FILE.unlink(missing_ok=True)
    except OSError:
        pass


def write_cache(resumes: dict[str, str]) -> None:
    try:
        CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
        CACHE_FILE.write_text(json.dumps({"at": time.time(), "resumes": resumes},
                                         ensure_ascii=False, indent=1),
                              encoding="utf-8")
    except OSError as exc:
        log.warning("[hh_resumes] кэш не записан: %s", exc)


# Карточка резюме на /applicant/resumes — это ссылка целиком: тип занятости
# перед названием, доход и формат работы после. Оставляем само название.
_CARD_HEAD_RE = re.compile(
    r"^(?:permanent job|temporary job|part[- ]time job|internship|project work|"
    r"постоянная работа|временная работа|частичная занятость|стажировка|"
    r"проектная работа)\s+", re.I)
_CARD_TAIL_RE = re.compile(
    r"\s+(?:income level|уровень дохода|доход не указан|зарплата не указана|"
    r"\d[\d\s ]*(?:₽|руб|\$|€)).*$", re.I)


def clean_title(text: str) -> str:
    t = " ".join((text or "").split())
    t = _CARD_HEAD_RE.sub("", t)
    return _CARD_TAIL_RE.sub("", t).strip() or t


async def account_resumes(page, use_cache: bool = True) -> dict[str, str]:
    """{hash: название} резюме аккаунта со страницы /applicant/resumes."""
    if use_cache:
        cached = read_cache()
        if cached:
            return {h: clean_title(t) for h, t in cached.items()}
    await page.goto("https://hh.ru/applicant/resumes", wait_until="domcontentloaded")
    found: dict[str, str] = {}
    for _ in range(12):                       # список дорисовывается скриптом
        for a in await page.locator("a[href*='/resume/']").all():
            try:
                href = await a.get_attribute("href") or ""
                text = " ".join((await a.inner_text()).split())
            except Exception:  # noqa: BLE001 — элемент перерисовался
                continue
            m = RESUME_HREF_RE.search(href)
            if m and len(text) > 3:
                text = clean_title(text)[:150]
                if len(text) > len(found.get(m.group(1), "")):
                    found[m.group(1)] = text
        if found:
            break
        await page.wait_for_timeout(1000)
    if found:
        write_cache(found)
    return found


def match_profiles(resumes: dict[str, str]) -> tuple[dict[str, dict], list[str]]:
    """{ключ профиля: {hash, title}} и список проблем сопоставления."""
    from jobsignal import profiles
    mapping: dict[str, dict] = {}
    problems: list[str] = []
    for p in profiles.all_profiles():
        if not p.hh_resume_title:
            problems.append(f"у профиля «{p.name}» не задано hh_resume_title")
            continue
        hits = [(h, t) for h, t in resumes.items()
                if p.hh_resume_title.lower() in t.lower()]
        if len(hits) == 1:
            mapping[p.key] = {"hash": hits[0][0], "title": hits[0][1],
                              "key": p.hh_resume_title}
        elif not hits:
            problems.append(f"профилю «{p.name}» не нашлось резюме на hh со словами "
                            f"«{p.hh_resume_title}»")
        else:
            problems.append(f"профилю «{p.name}» подходят {len(hits)} резюме "
                            f"(«{p.hh_resume_title}») — нужно ровно одно")
    return mapping, problems


def choice_needed() -> bool:
    """Выбирать резюме нужно, только если профилей больше одного."""
    from jobsignal import profiles
    return len(profiles.all_profiles()) > 1


# ── пауза после неверного резюме ─────────────────────────────────────────────

def blocked() -> dict | None:
    try:
        return json.loads(BLOCK_FILE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None


def block(reason: str, **extra) -> None:
    try:
        BLOCK_FILE.parent.mkdir(parents=True, exist_ok=True)
        BLOCK_FILE.write_text(json.dumps(
            {"reason": reason, "at": time.strftime("%d.%m.%Y %H:%M UTC", time.gmtime()),
             **extra}, ensure_ascii=False, indent=1), encoding="utf-8")
    except OSError as exc:
        log.error("[hh_resumes] паузу записать не удалось: %s", exc)


def unblock() -> bool:
    if BLOCK_FILE.exists():
        BLOCK_FILE.unlink()
        return True
    return False
