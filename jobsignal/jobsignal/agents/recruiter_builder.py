"""
RecruiterBuilder — собирает базу рекрутёров из уже распознанных вакансий.

Логика:
  1. Берёт все вакансии с recruiter_handle или у которых company заполнена
  2. Нормализует handle, извлекает имя из текста поста если есть
  3. Складывает в таблицу recruiters (дедупликация по handle)
  4. Для вакансий без handle — ищет рекрутёра по company

Таблица recruiters создаётся через raw SQL (старая схема БД).
"""
from __future__ import annotations

import logging
import re
import sqlite3
from datetime import datetime, timezone

log = logging.getLogger("jobsignal")

DB_PATH = "data/jobsignal.db"

# ── schema ────────────────────────────────────────────────────────────────────

CREATE_RECRUITERS = """
CREATE TABLE IF NOT EXISTS recruiters (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    handle VARCHAR(128),          -- @tg_username
    name VARCHAR(256),            -- имя если известно
    company VARCHAR(256),         -- компания
    source VARCHAR(64),           -- tg / hh / manual
    tg_handle VARCHAR(128),       -- дубль для явности
    linkedin_url VARCHAR(512),
    email VARCHAR(256),
    hh_profile_url VARCHAR(512),
    notes TEXT,
    vacancy_count INTEGER DEFAULT 1,
    last_seen_at DATETIME,
    created_at DATETIME,
    UNIQUE(handle)
)
"""

CREATE_RECRUITER_VACANCY = """
CREATE TABLE IF NOT EXISTS recruiter_vacancy (
    recruiter_id INTEGER NOT NULL,
    vacancy_id INTEGER NOT NULL,
    PRIMARY KEY (recruiter_id, vacancy_id)
)
"""


# ── helpers ───────────────────────────────────────────────────────────────────

_NAME_PATTERNS = [
    re.compile(r"(?:контакт|написать|обращаться|связаться)[:\s]+([А-ЯЁA-Z][а-яёa-z]+(?:\s[А-ЯЁA-Z][а-яёa-z]+)?)", re.IGNORECASE),
    re.compile(r"([А-ЯЁ][а-яё]+\s[А-ЯЁ][а-яё]+)\s*[\-–]\s*(?:HR|рекрутер|recruiter|хантер)", re.IGNORECASE),
    re.compile(r"(?:HR|рекрутер|recruiter)[:\s]+([А-ЯЁA-Z][а-яёa-z]+(?:\s[А-ЯЁA-Z][а-яёa-z]+)?)", re.IGNORECASE),
]

def _extract_name(text: str) -> str | None:
    if not text:
        return None
    for pat in _NAME_PATTERNS:
        m = pat.search(text)
        if m:
            name = m.group(1).strip()
            if 2 < len(name) < 60:
                return name
    return None


def _normalize_handle(h: str) -> str | None:
    if not h:
        return None
    h = h.strip().lstrip("@")
    # reject non-TG handles (LinkedIn, email, etc.)
    if re.search(r"[./: ]|linkedin|http", h, re.IGNORECASE):
        return None
    if not re.match(r"^[A-Za-z0-9_]{4,32}$", h):
        return None
    return f"@{h}"


def _normalize_company(c: str) -> str | None:
    if not c or c in ("—", "-", "None", "null"):
        return None
    return c.strip()


# ── main builder ──────────────────────────────────────────────────────────────

class RecruiterBuilder:
    def __init__(self, db_path: str = DB_PATH):
        self.db_path = db_path

    def _conn(self) -> sqlite3.Connection:
        return sqlite3.connect(self.db_path)

    def setup(self):
        """Create recruiter tables if not exist."""
        conn = self._conn()
        conn.execute(CREATE_RECRUITERS)
        conn.execute(CREATE_RECRUITER_VACANCY)
        conn.commit()
        conn.close()
        log.info("[recruiters] таблицы готовы")

    def build(self) -> dict:
        """Extract recruiters from existing vacancies."""
        self.setup()
        conn = self._conn()
        now = datetime.now(timezone.utc).isoformat()

        # fetch all vacancies with handle or company
        vacancies = conn.execute("""
            SELECT v.id, v.recruiter_handle, v.company, v.link,
                   rp.raw_text, rp.text
            FROM vacancies v
            LEFT JOIN raw_posts rp ON rp.id = v.raw_post_id
            WHERE v.recruiter_handle IS NOT NULL
               OR v.company IS NOT NULL
        """).fetchall()

        added = 0
        updated = 0
        linked = 0

        for vac_id, handle_raw, company_raw, link, raw_text, text in vacancies:
            handle = _normalize_handle(handle_raw or "")
            company = _normalize_company(company_raw)
            post_text = raw_text or text or ""
            name = _extract_name(post_text)

            # determine source
            source = "hh" if link and "hh.ru" in link else "tg"

            if handle:
                # upsert by handle
                existing = conn.execute(
                    "SELECT id, vacancy_count, company FROM recruiters WHERE handle=?",
                    (handle,)
                ).fetchone()

                if existing:
                    rec_id, count, existing_company = existing
                    # update company if we now know it and didn't before
                    new_company = company or existing_company
                    conn.execute("""
                        UPDATE recruiters
                        SET vacancy_count=?, last_seen_at=?, company=COALESCE(?, company)
                        WHERE id=?
                    """, (count + 1, now, new_company, rec_id))
                    updated += 1
                else:
                    conn.execute("""
                        INSERT OR IGNORE INTO recruiters
                        (handle, name, company, source, tg_handle, last_seen_at, created_at)
                        VALUES (?,?,?,?,?,?,?)
                    """, (handle, name, company, source, handle, now, now))
                    rec_id = conn.execute(
                        "SELECT id FROM recruiters WHERE handle=?", (handle,)
                    ).fetchone()
                    if rec_id:
                        rec_id = rec_id[0]
                        added += 1

                # link recruiter <-> vacancy
                if rec_id:
                    try:
                        conn.execute(
                            "INSERT OR IGNORE INTO recruiter_vacancy (recruiter_id, vacancy_id) VALUES (?,?)",
                            (rec_id, vac_id)
                        )
                        linked += 1
                    except Exception:
                        pass

            elif company:
                # no handle but has company — we'll find recruiter by company later
                pass

        conn.commit()
        conn.close()

        result = {
            "agent": "recruiter_builder",
            "added": added,
            "updated": updated,
            "linked": linked,
        }
        log.info("[recruiters] добавлено: %d, обновлено: %d, связей: %d", added, updated, linked)
        return result

    def find_for_company(self, company: str) -> list[dict]:
        """Find recruiters by company name (fuzzy)."""
        if not company:
            return []
        conn = self._conn()
        rows = conn.execute("""
            SELECT handle, name, company, tg_handle, linkedin_url, email,
                   hh_profile_url, vacancy_count, last_seen_at
            FROM recruiters
            WHERE company LIKE ?
            ORDER BY vacancy_count DESC
            LIMIT 5
        """, (f"%{company}%",)).fetchall()
        conn.close()
        return [
            {
                "handle": r[0], "name": r[1], "company": r[2],
                "tg_handle": r[3], "linkedin_url": r[4], "email": r[5],
                "hh_profile_url": r[6], "vacancy_count": r[7],
                "last_seen_at": r[8],
            }
            for r in rows
        ]

    def all_recruiters(self, limit: int = 200) -> list[dict]:
        """List all known recruiters."""
        conn = self._conn()
        rows = conn.execute("""
            SELECT handle, name, company, source, tg_handle,
                   linkedin_url, email, vacancy_count, last_seen_at
            FROM recruiters
            ORDER BY vacancy_count DESC, last_seen_at DESC
            LIMIT ?
        """, (limit,)).fetchall()
        conn.close()
        return [
            {
                "handle": r[0], "name": r[1], "company": r[2],
                "source": r[3], "tg_handle": r[4], "linkedin_url": r[5],
                "email": r[6], "vacancy_count": r[7], "last_seen_at": r[8],
            }
            for r in rows
        ]
