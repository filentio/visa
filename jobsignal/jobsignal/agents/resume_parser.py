"""
ResumeParser — парсит PDF резюме и сохраняет текст в БД.

Запускается автоматически при загрузке резюме через /resumes.
Текст резюме используется при генерации сопроводительных писем
вместо хардкода — исключает галлюцинации.
"""
from __future__ import annotations

import logging
import sqlite3
from pathlib import Path

log = logging.getLogger("jobsignal")

DB_PATH = "data/jobsignal.db"
RESUME_DIR = Path("config/resumes")

PROFILE_MAP = {
    "ai_pm": "Senior AI PM",
    "cpo": "CPO / Head of Product",
    "pm": "Senior PM/PO",
}


def _ensure_table():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS resumes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            profile_key VARCHAR(32) UNIQUE NOT NULL,
            profile_name VARCHAR(128),
            filename VARCHAR(256),
            raw_text TEXT,
            updated_at DATETIME DEFAULT (datetime('now'))
        )
    """)
    conn.commit()
    conn.close()


def _extract_pdf_text(pdf_path: str) -> str:
    """Extract text from PDF using pypdf (already in venv)."""
    try:
        import pypdf
        reader = pypdf.PdfReader(pdf_path)
        pages = []
        for page in reader.pages:
            text = page.extract_text()
            if text:
                pages.append(text.strip())
        return "\n\n".join(pages)
    except Exception as e:
        log.error("[resume_parser] pypdf error for %s: %s", pdf_path, e)
        return ""


def _extract_pdf_text_fallback(pdf_path: str) -> str:
    """Fallback: pdfminer if available."""
    try:
        from pdfminer.high_level import extract_text
        return extract_text(pdf_path)
    except Exception:
        pass
    # last resort: pdfplumber
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            return "\n\n".join(
                p.extract_text() or "" for p in pdf.pages
            )
    except Exception:
        return ""


def parse_and_save(profile_key: str) -> dict:
    """Parse PDF for given profile_key and save to DB."""
    _ensure_table()

    pdf_path = RESUME_DIR / f"{profile_key}.pdf"
    if not pdf_path.exists():
        return {"ok": False, "error": f"PDF not found: {pdf_path}"}

    text = _extract_pdf_text(str(pdf_path))
    if not text:
        text = _extract_pdf_text_fallback(str(pdf_path))
    if not text:
        return {"ok": False, "error": "could not extract text from PDF"}

    profile_name = PROFILE_MAP.get(profile_key, profile_key)
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        INSERT INTO resumes (profile_key, profile_name, filename, raw_text, updated_at)
        VALUES (?, ?, ?, ?, datetime('now'))
        ON CONFLICT(profile_key) DO UPDATE SET
            profile_name=excluded.profile_name,
            filename=excluded.filename,
            raw_text=excluded.raw_text,
            updated_at=excluded.updated_at
    """, (profile_key, profile_name, pdf_path.name, text))
    conn.commit()
    conn.close()

    log.info("[resume_parser] %s: %d chars saved", profile_key, len(text))
    return {"ok": True, "profile_key": profile_key, "chars": len(text)}


def get_resume_text(profile_key: str) -> str | None:
    """Get resume text from DB for given profile_key."""
    _ensure_table()
    conn = sqlite3.connect(DB_PATH)
    row = conn.execute(
        "SELECT raw_text FROM resumes WHERE profile_key=?", (profile_key,)
    ).fetchone()
    conn.close()
    return row[0] if row else None


def parse_all() -> dict:
    """Parse all available resume PDFs."""
    _ensure_table()
    results = {}
    for key in PROFILE_MAP:
        pdf = RESUME_DIR / f"{key}.pdf"
        if pdf.exists():
            results[key] = parse_and_save(key)
        else:
            results[key] = {"ok": False, "error": "no PDF"}
    return results
