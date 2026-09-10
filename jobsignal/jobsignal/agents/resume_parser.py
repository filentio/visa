"""
ResumeParser — парсит PDF резюме и сохраняет текст в БД.

Запускается автоматически при загрузке резюме через /resumes.
Текст резюме используется при генерации сопроводительных писем
вместо хардкода — исключает галлюцинации.

Какой PDF брать для профиля, решает pdf_path_for(): отдельных резюме по
профилям больше нет, все три указывают в profiles.yaml на общее
config/master_cv.pdf.
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


def pdf_path_for(profile_key: str) -> Path | None:
    """PDF профиля: свой файл в config/resumes/, иначе resume_path из profiles.yaml.

    Позиционирование единое, поэтому все три профиля указывают на общее
    config/master_cv.pdf — обновлять надо один файл. Личный PDF профиля, если
    его загрузили через /resumes, остаётся главнее: иначе загрузка молча
    уходила бы в никуда.
    """
    own = RESUME_DIR / f"{profile_key}.pdf"
    if own.exists():
        return own

    profiles_yaml = Path("config/profiles.yaml")
    if not profiles_yaml.exists():
        return None
    import yaml
    with profiles_yaml.open(encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    name = PROFILE_MAP.get(profile_key)
    for p in data.get("profiles", []):
        if p.get("name") != name:
            continue
        rp = p.get("resume_path")
        if not rp:
            return None
        path = Path(rp)
        return path if path.exists() else None
    return None


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
    except ImportError:
        pass  # библиотеки нет — штатно пробуем следующую
    except Exception as exc:
        log.warning("[resume] pdfminer не смог %s: %s", pdf_path, exc)
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

    pdf_path = pdf_path_for(profile_key)
    if pdf_path is None:
        return {"ok": False,
                "error": f"PDF не найден для профиля {profile_key}: нет ни "
                         f"{RESUME_DIR}/{profile_key}.pdf, ни resume_path из profiles.yaml"}

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
