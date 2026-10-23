"""Карьерные профили кандидата — единственный источник правды.

Зачем. Три профиля были прописаны в коде в дюжине мест: подпись письма в
боте, словарь «имя → ключ файла резюме» в боте, дашборде и разборщике
резюме, умолчание «Senior PM/PO» в автоотклике, композиторе и дашборде,
список на странице резюме, поисковые запросы hh. 23.10 кандидат сменил
позиционирование на два новых резюме, и поменять это можно было только
правкой кода в каждом из этих мест. Для второго клиента — тем более.

Теперь всё о профиле живёт в config/profiles.yaml:

    profiles:
      - key: po                         # латиница: имя файла резюме, ключ в базе резюме
        name: "Product Owner (финтех, AI)"   # имя в оценках (match_scores.profile_key)
        label: "Product Owner / Менеджер продукта"   # как назвать в письме
        resume_path: ./config/resumes/po.pdf   # PDF: прикладывается и даёт текст для писем
        cv_path: ./config/cv_po.md      # необязательно: выжимка для оценки; нет — текст PDF
        hh_resume_title: "Product Owner"   # часть названия резюме на hh.ru для выбора в отклике
        focus: "…"                      # одной строкой: на какие роли целимся (оценка, разведка)
        searches: ["Product Owner", …]  # запросы поиска hh

Порядок профилей важен: первый — профиль по умолчанию.
"""
from __future__ import annotations

import logging
from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path

import yaml

log = logging.getLogger("jobsignal")

ROOT = Path(__file__).resolve().parent.parent
PROFILES_YAML = ROOT / "config" / "profiles.yaml"
RESUME_DIR = ROOT / "config" / "resumes"


@dataclass(frozen=True)
class Profile:
    key: str
    name: str
    label: str = ""
    resume_path: str = ""
    cv_path: str = ""
    hh_resume_title: str = ""
    focus: str = ""
    searches: tuple = field(default_factory=tuple)

    @property
    def letter_label(self) -> str:
        return self.label or self.name

    def resume_pdf(self) -> Path | None:
        """PDF профиля: загруженный в кабинете config/resumes/<key>.pdf главнее
        resume_path — иначе загрузка через /resumes молча уходила бы в никуда."""
        own = RESUME_DIR / f"{self.key}.pdf"
        if own.exists():
            return own
        if self.resume_path:
            p = _abs(self.resume_path)
            if p.exists():
                return p
        return None

    def cv_text(self) -> str:
        """Текст для оценки вакансий: выжимка cv_path, иначе текст PDF."""
        if self.cv_path:
            p = _abs(self.cv_path)
            if p.exists():
                text = p.read_text(encoding="utf-8").strip()
                if text:
                    return text
        pdf = self.resume_pdf()
        if pdf is not None:
            from jobsignal.agents.resume_parser import extract_text
            return extract_text(str(pdf)).strip()
        return ""


def _abs(path: str) -> Path:
    p = Path(path)
    return p if p.is_absolute() else ROOT / p


@lru_cache(maxsize=1)
def _load(mtime: float) -> tuple[Profile, ...]:
    if not PROFILES_YAML.exists():
        log.warning("[profiles] нет %s — профилей нет", PROFILES_YAML)
        return ()
    data = yaml.safe_load(PROFILES_YAML.read_text(encoding="utf-8")) or {}
    out = []
    for i, p in enumerate(data.get("profiles") or []):
        name = (p.get("name") or "").strip()
        if not name:
            continue
        key = (p.get("key") or f"p{i + 1}").strip()
        searches = p.get("searches") or []
        out.append(Profile(
            key=key, name=name, label=(p.get("label") or "").strip(),
            resume_path=p.get("resume_path") or "", cv_path=p.get("cv_path") or "",
            hh_resume_title=(p.get("hh_resume_title") or "").strip(),
            focus=(p.get("focus") or "").strip(),
            searches=tuple(str(s).strip() for s in searches if str(s).strip()),
        ))
    return tuple(out)


def all_profiles() -> tuple[Profile, ...]:
    mtime = PROFILES_YAML.stat().st_mtime if PROFILES_YAML.exists() else 0.0
    return _load(mtime)


def names() -> list[str]:
    return [p.name for p in all_profiles()]


def keys() -> list[str]:
    return [p.key for p in all_profiles()]


def default() -> Profile | None:
    ps = all_profiles()
    return ps[0] if ps else None


def default_name() -> str:
    d = default()
    return d.name if d else ""


def by_name(name: str | None) -> Profile | None:
    for p in all_profiles():
        if p.name == name:
            return p
    return None


def by_key(key: str | None) -> Profile | None:
    for p in all_profiles():
        if p.key == key:
            return p
    return None


def resolve(name_or_key: str | None) -> Profile | None:
    """Профиль по имени или ключу; неизвестный (старый, снятый) — по умолчанию."""
    return by_name(name_or_key) or by_key(name_or_key) or default()


def focus_text() -> str:
    """На какие роли целимся — для оценки каналов и подсказок модели."""
    parts = [f"{p.name}: {p.focus}" if p.focus else p.name for p in all_profiles()]
    return "; ".join(parts)
