"""
fix_contacts.py — одноразовый скрипт для исправления уже сохранённых вакансий:
  1. Убирает невалидные TG-хэндлы (LinkedIn URL, email и т.д.)
  2. Заменяет плейсхолдерные ссылки [ссылка] на null / post_url
  3. Обновляет contact_type
Запуск: python fix_contacts.py
"""
import sys
import re
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from dotenv import load_dotenv
load_dotenv("config/.env")

from jobsignal.db import get_session_factory, Vacancy, RawPost, Channel
from jobsignal.agents.parser import _normalize_handle, _is_real_link, _extract_links_from_text, _pick_best_link
from jobsignal.agents.outreach import detect_contact_type

sf = get_session_factory()
session = sf()

vacancies = session.query(Vacancy).all()
fixed_handles = 0
fixed_links = 0
fixed_contact_type = 0

for v in vacancies:
    changed = False

    # ── fix recruiter_handle ──────────────────────────────────────────────────
    if v.recruiter_handle:
        clean = _normalize_handle(v.recruiter_handle)
        if clean != v.recruiter_handle:
            print(f"  handle fix [{v.id}] {v.recruiter_handle!r} → {clean!r}")
            # if it was a real link, move to vacancy link if link is empty
            if not clean and _is_real_link(v.recruiter_handle) and not v.link:
                v.link = v.recruiter_handle
                fixed_links += 1
            v.recruiter_handle = clean
            fixed_handles += 1
            changed = True

    # ── fix link ──────────────────────────────────────────────────────────────
    if v.link and not _is_real_link(v.link):
        # try to recover from raw post text
        post = session.query(RawPost).filter_by(id=v.raw_post_id).first()
        post_url = post.post_url if post else None
        text_links = _extract_links_from_text(post.text or "") if post else []
        best = _pick_best_link(text_links, post_url)
        print(f"  link fix [{v.id}] {v.link!r} → {best!r}")
        v.link = best
        fixed_links += 1
        changed = True
    elif not v.link:
        # Try to set post_url as fallback link
        post = session.query(RawPost).filter_by(id=v.raw_post_id).first()
        if post and post.post_url:
            v.link = post.post_url
            fixed_links += 1
            changed = True

    # ── update contact_type ───────────────────────────────────────────────────
    new_ct = detect_contact_type(v)
    if v.contact_type != new_ct:
        v.contact_type = new_ct
        fixed_contact_type += 1
        changed = True

try:
    session.commit()
    print(f"\nДоне:")
    print(f"  Исправлено хэндлов:      {fixed_handles}")
    print(f"  Исправлено ссылок:       {fixed_links}")
    print(f"  Обновлено contact_type:  {fixed_contact_type}")
except Exception as e:
    session.rollback()
    print(f"Ошибка: {e}")
finally:
    session.close()
