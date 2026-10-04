"""Сколько в телеграме живых контактов, а сколько каналов.

Показывает, что даст отбор по частоте (jobsignal/tg_contacts.py), не меняя
ничего в базе. Нужен, чтобы выбрать порог осознанно и увидеть, кого именно
мы отсекаем — ошибка здесь стоит пропущенных вакансий.

    cd /opt/jobsignal_local && .venv/bin/python tools/tg_pool.py
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from dotenv import load_dotenv  # noqa: E402
load_dotenv("config/.env")

from sqlalchemy import func, select  # noqa: E402

from jobsignal.db import Vacancy, VacancyStatus, get_session_factory  # noqa: E402
from jobsignal.tg_contacts import MASS_MIN, mass_handles  # noqa: E402


def main() -> int:
    s = get_session_factory()()
    mass = mass_handles(s)
    print(f"порог «не человек»: от {MASS_MIN} вакансий на контакт")
    print(f"контактов признано каналами: {len(mass)}\n")

    rows = s.execute(
        select(Vacancy.recruiter_handle, func.count())
        .where(Vacancy.recruiter_handle.isnot(None),
               Vacancy.contact_type == "tg")
        .group_by(Vacancy.recruiter_handle)
        .order_by(func.count().desc())
    ).all()

    cut = [(h, n) for h, n in rows if h in mass]
    keep = [(h, n) for h, n in rows if h not in mass]
    print(f"ОТСЕКАЕМ: {len(cut)} контактов, {sum(n for _, n in cut)} вакансий")
    for h, n in cut[:12]:
        print(f"    {h}: {n}")
    print(f"\nОСТАВЛЯЕМ: {len(keep)} контактов, {sum(n for _, n in keep)} вакансий")
    print("    примеры с краю порога:")
    for h, n in keep[:8]:
        print(f"    {h}: {n}")

    # Сколько из оставшихся реально дошло бы до отклика: прошли порог оценки
    # и ещё не обработаны.
    live = s.execute(
        select(func.count()).select_from(Vacancy).where(
            Vacancy.contact_type == "tg",
            Vacancy.recruiter_handle.isnot(None),
            Vacancy.recruiter_handle.notin_(mass) if mass else True,
            Vacancy.status == VacancyStatus.matched,
        )
    ).scalar()
    print(f"\nиз них уже прошли порог оценки и ждут отклика: {live}")
    print("(остальные телеграм-вакансии сейчас не оцениваются — см. MATCH_SOURCES)")
    s.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
