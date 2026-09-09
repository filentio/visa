"""RawDeduplicator — грубый дедуп сырых постов ДО парсера.

Одна и та же вакансия висит в нескольких каналах, и до сих пор каждый репост
уходил в модель на разбор: DeduplicatorAgent выбрасывал лишние уже после, когда
за них было заплачено. Здесь дубли ловятся по хешу нормализованного текста —
до единого вызова модели. Каждый отсеянный репост экономит вызов целиком, а не
удешевляет его.

Точный дедуп по роли и компании остаётся в dedup.py: эти поля появляются только
после разбора, до парсера их взять негде.

Дубль помечается parsed=True (парсер берёт только parsed=False) и вакансии не
создаёт — вакансия уже есть от оригинала.
"""
from __future__ import annotations

import hashlib
import logging

from sqlalchemy import select

from ..db import RawPost, get_session_factory
from .base import BaseAgent
from .dedup import _normalize

log = logging.getLogger("jobsignal")

# столько же пропускает парсер (_parse_post) — считать их дублями смысла нет
MIN_TEXT_LEN = 30


def text_signature(text: str) -> str:
    """Ключ дубля: хеш нормализованного текста поста.

    Нормализация — та же, что у точного дедупа (регистр, пунктуация, пробелы).
    Ссылки и @handle намеренно НЕ вырезаются: на текущей базе это ловит всего
    +90 постов, но склеивает посты, состоящие из одной ссылки, — а это разные
    вакансии.
    """
    return hashlib.sha1(_normalize(text).encode("utf-8")).hexdigest()[:16]


class RawDeduplicatorAgent(BaseAgent):
    name = "raw_deduplicator"

    def run(self) -> dict:
        Session = get_session_factory()
        with Session() as s:
            posts = s.execute(select(RawPost).order_by(RawPost.id)).scalars().all()

            # Хеши уже разобранных постов тоже идут в set: репост приходит
            # через день-два, когда оригинал давно parsed — иначе такой дубль
            # не поймать. Пересчёт по всей базе стоит миллисекунды.
            seen: set[str] = set()
            for p in posts:
                text = (p.text or "").strip()
                if len(text) < MIN_TEXT_LEN:
                    continue
                if p.parsed:
                    seen.add(text_signature(text))

            dups = 0
            considered = 0
            for p in posts:
                if p.parsed:
                    continue
                text = (p.text or "").strip()
                if len(text) < MIN_TEXT_LEN:
                    continue
                considered += 1
                sig = text_signature(text)
                if sig in seen:
                    log.debug("[raw_dedup] пост %d — дубль по хешу %s, разбор пропущен",
                              p.id, sig)
                    p.parsed = True
                    dups += 1
                    continue
                seen.add(sig)
            s.commit()

        log.info("[raw_dedup] к разбору было: %d, отсеяно дублей: %d, "
                 "останется вызовов парсера: %d", considered, dups, considered - dups)
        return {"agent": self.name, "considered": considered,
                "duplicates": dups, "to_parse": considered - dups}
