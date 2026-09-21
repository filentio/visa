"""Переоценка вакансий после перехода на полный текст — со сравнением.

До 21.09 матчер читал `Vacancy.description` — пересказ парсера в одну фразу
(184 знака у hh против 2854 в исходном посте). Значит все накопленные баллы
посчитаны не по требованиям вакансии. Этот скрипт переоценивает выборку и
печатает старый балл рядом с новым, чтобы решение о полной переоценке
опиралось на числа.

    .venv/bin/python tools/rescore.py --limit 20            # посмотреть, кого возьмёт
    .venv/bin/python tools/rescore.py --limit 20 --apply    # переоценить

Без --apply ничего не меняет. Старые баллы перед заменой сохраняются в
/tmp/rescore_backup_<время>.json — восстановить вручную, если понадобится.

Берём только то, что имеет смысл переоценивать: основные вакансии, по которым
отклик ещё не ушёл. Историю переоценивать незачем — вакансии закрылись, а
расход реальный.
"""
from __future__ import annotations

import argparse
import json
import logging
import sys
import time
from pathlib import Path

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s INFO jobsignal: %(message)s",
                    datefmt="%H:%M:%S")

from dotenv import load_dotenv  # noqa: E402
load_dotenv("config/.env")

from sqlalchemy import select  # noqa: E402
from sqlalchemy.orm import selectinload  # noqa: E402

from jobsignal.db import (Application, MatchScore, Vacancy,  # noqa: E402
                          VacancyStatus, get_session_factory)


def pick(session, limit: int, days: int, source: str | None):
    """Вакансии, которые стоит переоценить: свежие, основные, без отклика."""
    q = (select(Vacancy)
         .options(selectinload(Vacancy.raw_post),
                  selectinload(Vacancy.match_scores))
         .where(Vacancy.is_primary.is_(True),
                Vacancy.created_at >= _cutoff(days),
                ~Vacancy.applications.any(Application.sent_at.isnot(None))))
    if source:
        q = q.where(Vacancy.contact_type == source)
    rows = session.execute(q).scalars().all()
    # Сначала те, у кого исходный текст есть и он длиннее пересказа: именно у
    # них оценка и менялась. Вакансия без поста переоценку не изменит.
    rows.sort(key=lambda v: -(len((v.raw_post.text if v.raw_post else "") or "")))
    return rows[:limit]


def _cutoff(days: int):
    from datetime import datetime, timedelta, timezone
    return datetime.now(timezone.utc) - timedelta(days=days)


def best(v: Vacancy) -> int:
    return max((m.score for m in v.match_scores), default=-1)


def main() -> int:
    ap = argparse.ArgumentParser(prog="tools/rescore.py")
    ap.add_argument("--limit", type=int, default=20)
    ap.add_argument("--days", type=int, default=7,
                    help="насколько свежие вакансии брать (по умолчанию неделя)")
    ap.add_argument("--source", default=None,
                    help="contact_type: hh / tg / form / linkedin")
    ap.add_argument("--apply", action="store_true",
                    help="без него только показывает выборку")
    opts = ap.parse_args()

    sf = get_session_factory()
    session = sf()
    vacs = pick(session, opts.limit, opts.days, opts.source)
    if not vacs:
        print("подходящих вакансий нет — проверь --days и --source")
        return 1

    print(f"{'id':>6} {'балл':>5} {'пересказ':>9} {'полный текст':>13}  роль")
    for v in vacs:
        full = len((v.raw_post.text if v.raw_post else "") or "")
        print(f"{v.id:>6} {best(v):>5} {len(v.description or ''):>9} {full:>13}  "
              f"{(v.role or '?')[:40]}")

    if not opts.apply:
        print(f"\nвыбрано {len(vacs)} вакансий. Это просмотр; для переоценки "
              f"добавь --apply")
        return 0

    # Старые баллы на диск — до того, как удалим строки.
    backup = Path(f"/tmp/rescore_backup_{int(time.time())}.json")
    old = {v.id: {"best": best(v), "status": v.status.value,
                  "scores": [{"profile": m.profile_key, "score": m.score,
                              "reason": m.reason} for m in v.match_scores]}
           for v in vacs}
    backup.write_text(json.dumps(old, ensure_ascii=False, indent=1))
    print(f"\nстарые баллы сохранены: {backup}")

    ids = [v.id for v in vacs]
    for v in vacs:
        for m in list(v.match_scores):
            session.delete(m)
        # Матчер берёт только NEW — иначе выборку он просто не увидит.
        v.status = VacancyStatus.new
    session.commit()
    session.close()

    from jobsignal.agents.matcher import MatcherAgent
    from jobsignal.config import get_config
    result = MatcherAgent(get_config()).run(limit=len(ids))
    print(f"матчер: {result}")

    session = sf()
    print(f"\n{'id':>6} {'было':>5} {'стало':>6} {'дельта':>7}  роль")
    moved = 0
    for vid in ids:
        v = session.get(Vacancy, vid)
        was, now = old[vid]["best"], best(v)
        delta = now - was
        if abs(delta) >= 10:
            moved += 1
        mark = "  ←" if abs(delta) >= 20 else ""
        print(f"{vid:>6} {was:>5} {now:>6} {delta:>+7}{mark}  {(v.role or '?')[:36]}")
    print(f"\nсдвинулось на 10+ баллов: {moved} из {len(ids)}")
    print("стрелкой отмечены сдвиги на 20 и больше")
    session.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
