"""JobSignal dashboard — Flask app."""
from __future__ import annotations

import base64
import json
import logging
import os
from datetime import datetime, timezone, timedelta
from functools import wraps
from pathlib import Path

from flask import (
    Flask, render_template, jsonify, request,
    send_file, abort, Response, redirect, url_for
)
from sqlalchemy import desc, func

from jobsignal.db import (
    get_session_factory, Vacancy, VacancyStatus, MatchScore,
    Application, Channel, ChannelCandidate
)
from jobsignal.agents.outreach import detect_contact_type, get_outreach_data
from jobsignal.agents.composer import Composer, STYLES

log = logging.getLogger("jobsignal")

app = Flask(__name__)

_sf = None

def _session():
    global _sf
    if _sf is None:
        _sf = get_session_factory()
    return _sf()


# ── auth ──────────────────────────────────────────────────────────────────────

def _check_auth(username: str, password: str) -> bool:
    u = os.environ.get("DASHBOARD_USER", "admin")
    p = os.environ.get("DASHBOARD_PASS", "")
    if not p:
        return True
    return username == u and password == p


def _requires_auth(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        auth = request.authorization
        if not auth or not _check_auth(auth.username, auth.password):
            return Response(
                "Нужна авторизация", 401,
                {"WWW-Authenticate": 'Basic realm="JobSignal"'}
            )
        return f(*args, **kwargs)
    return decorated


app.before_request(lambda: None)  # placeholder; auth applied per-route


# ── helpers ───────────────────────────────────────────────────────────────────

def _best_score(vacancy) -> tuple[int, str]:
    if not vacancy.match_scores:
        return 0, ""
    best = max(vacancy.match_scores, key=lambda s: s.score)
    return best.score, best.profile_key


def _rate_status() -> dict:
    session = _session()
    now = datetime.now(timezone.utc)
    per_hour = int(os.environ.get("OUTREACH_PER_HOUR", "3"))
    per_day = int(os.environ.get("OUTREACH_PER_DAY", "15"))

    hour_ago = now - timedelta(hours=1)
    day_ago = now - timedelta(days=1)

    sent_hour = (
        session.query(func.count(Application.id))
        .filter(Application.sent_at >= hour_ago)
        .scalar()
    )
    sent_day = (
        session.query(func.count(Application.id))
        .filter(Application.sent_at >= day_ago)
        .scalar()
    )
    session.close()

    allowed = min(per_hour - sent_hour, per_day - sent_day)
    allowed = max(0, allowed)

    next_slot_str = ""
    if allowed == 0:
        next_slot_str = (hour_ago + timedelta(hours=1)).strftime("%H:%M")

    return {
        "sent_hour": sent_hour,
        "sent_day": sent_day,
        "per_hour": per_hour,
        "per_day": per_day,
        "allowed_now": allowed,
        "next_slot": next_slot_str,
    }


def _is_new(v: Vacancy) -> bool:
    if not v.created_at:
        return False
    cutoff = datetime.now(timezone.utc) - timedelta(hours=24)
    created = v.created_at
    if created.tzinfo is None:
        created = created.replace(tzinfo=timezone.utc)
    return created >= cutoff


def _vacancy_to_dict(v: Vacancy) -> dict:
    best_score, best_profile = _best_score(v)
    scores = {s.profile_key: {"score": s.score, "reason": s.reason}
              for s in (v.match_scores or [])}
    contact_type = v.contact_type or detect_contact_type(v)
    return {
        "id": v.id,
        "role": v.role or "—",
        "company": v.company or "—",
        "recruiter_handle": v.recruiter_handle,
        "salary": v.salary or "—",
        "location": v.location or "—",
        "link": v.link or "",
        "description": v.description or "",
        "status": v.status.value,
        "created_at": v.created_at.strftime("%d.%m.%Y") if v.created_at else "—",
        "best_score": best_score,
        "best_profile": best_profile,
        "scores": scores,
        "draft_text": v.draft_text or "",
        "contact_type": contact_type,
        "has_tg": bool(v.recruiter_handle),
        "has_link": bool(v.link),
        "is_new": _is_new(v),
    }


# ── main page ─────────────────────────────────────────────────────────────────

@app.route("/")
@_requires_auth
def index():
    import yaml
    profiles = {}
    try:
        with open("config/profiles.yaml") as f:
            data = yaml.safe_load(f)
        for p in data.get("profiles", []):
            profiles[p["key"]] = p["name"]
    except Exception:
        profiles = {"ai_pm": "Senior AI PM", "cpo": "CPO / Head of Product", "pm": "Senior PM/PO"}
    return render_template("index.html", styles=list(STYLES.keys()), profiles=profiles)


@app.route("/api/vacancies")
@_requires_auth
def api_vacancies():
    session = _session()
    q = session.query(Vacancy).filter(Vacancy.is_primary == True)

    profile = request.args.get("profile", "")
    threshold = int(request.args.get("threshold", os.environ.get("MATCH_THRESHOLD", "70")))
    channel_id = request.args.get("channel_id", "")
    days = request.args.get("days", "")
    status_filter = request.args.get("status", "")
    search = request.args.get("search", "").strip().lower()
    queue_only = request.args.get("queue", "") == "1"
    contact_type_filter = request.args.get("contact_type", "")

    if days:
        cutoff = datetime.now(timezone.utc) - timedelta(days=int(days))
        q = q.filter(Vacancy.created_at >= cutoff)
    if channel_id:
        q = q.filter(Vacancy.channel_id == int(channel_id))
    if status_filter == "working":
        q = q.filter(Vacancy.status.in_([VacancyStatus.matched, VacancyStatus.drafted]))
    elif status_filter == "applied":
        q = q.filter(Vacancy.status == VacancyStatus.applied)
    elif status_filter == "replied":
        q = q.filter(Vacancy.status == VacancyStatus.replied)

    vacancies = q.order_by(desc(Vacancy.created_at)).all()

    result = []
    for v in vacancies:
        d = _vacancy_to_dict(v)

        # score filter
        if profile:
            score = d["scores"].get(profile, {}).get("score", 0)
        else:
            score = d["best_score"]
        if score < threshold:
            continue

        # queue: only matched/drafted with any contact
        if queue_only:
            if v.status not in (VacancyStatus.matched, VacancyStatus.drafted):
                continue
            if not v.recruiter_handle and not v.link:
                continue

        # contact type filter
        if contact_type_filter:
            if d["contact_type"] != contact_type_filter:
                continue

        if search:
            haystack = f"{d['role']} {d['company']}".lower()
            if search not in haystack:
                continue

        d["display_score"] = score
        result.append(d)

    result.sort(key=lambda x: x["display_score"], reverse=True)

    # channels for filter dropdown
    channels = [
        {"id": c.id, "title": c.title or c.handle}
        for c in session.query(Channel).filter_by(active=True).all()
    ]
    session.close()
    return jsonify({"vacancies": result, "channels": channels})


@app.route("/api/stats")
@_requires_auth
def api_stats():
    session = _session()
    total = session.query(func.count(Vacancy.id)).filter(
        Vacancy.is_primary == True,
        Vacancy.status.in_([VacancyStatus.matched, VacancyStatus.drafted,
                             VacancyStatus.applied, VacancyStatus.replied])
    ).scalar()
    working = session.query(func.count(Vacancy.id)).filter(
        Vacancy.status.in_([VacancyStatus.matched, VacancyStatus.drafted])
    ).scalar()
    applied = session.query(func.count(Application.id)).scalar()
    replied = session.query(func.count(Application.id)).filter(
        Application.replied_at.isnot(None)
    ).scalar()
    session.close()
    rate = _rate_status()
    # count pending channel candidates
    s2 = _session()
    new_candidates = s2.query(func.count(ChannelCandidate.id)).filter_by(status="pending").scalar()
    s2.close()
    return jsonify({
        "total": total, "working": working,
        "applied": applied, "replied": replied,
        "new_candidates": new_candidates or 0,
        **rate,
    })


@app.route("/api/generate_draft", methods=["POST"])
@_requires_auth
def api_generate_draft():
    data = request.get_json()
    vid = data.get("vacancy_id")
    style = data.get("style", "metric_hook")
    session = _session()
    v = session.get(Vacancy, vid)
    if not v:
        session.close()
        return jsonify({"error": "not found"}), 404
    try:
        composer = Composer()
        text = composer.generate(v, style=style)
        if not text or text.startswith("✗"):
            raise ValueError(text or "пустой ответ модели")
        v.draft_text = text
        if v.status == VacancyStatus.matched:
            v.status = VacancyStatus.drafted
        session.commit()
        session.close()
        return jsonify({"draft": text})
    except Exception as exc:
        session.rollback()
        session.close()
        log.error("generate_draft error vid=%s: %s", vid, exc)
        return jsonify({"error": str(exc)}), 500


@app.route("/api/outreach_data/<int:vid>")
@_requires_auth
def api_outreach_data(vid):
    """Returns outreach info for non-TG vacancies."""
    session = _session()
    v = session.get(Vacancy, vid)
    if not v:
        session.close()
        return jsonify({"error": "not found"}), 404
    _, best_profile = _best_score(v)
    profile_key = best_profile or "pm"
    # map profile key to short key
    key_map = {
        "Senior AI PM": "ai_pm",
        "CPO / Head of Product": "cpo",
        "Senior PM/PO": "pm",
    }
    short_key = key_map.get(profile_key, "pm")
    data = get_outreach_data(v, short_key)
    session.close()
    return jsonify(data)


@app.route("/api/mark_applied/<int:vid>", methods=["POST"])
@_requires_auth
def api_mark_applied(vid):
    rate = _rate_status()
    if rate["allowed_now"] <= 0:
        return jsonify({
            "error": "rate_limit",
            "next_slot": rate["next_slot"],
        }), 429

    session = _session()
    v = session.get(Vacancy, vid)
    if not v:
        session.close()
        return jsonify({"error": "not found"}), 404
    v.status = VacancyStatus.applied
    app_rec = Application(
        vacancy_id=vid,
        draft_text=v.draft_text,
        sent_at=datetime.now(timezone.utc),
    )
    session.add(app_rec)
    session.commit()
    session.close()
    return jsonify({"ok": True, "rate": _rate_status()})


@app.route("/api/mark_replied/<int:vid>", methods=["POST"])
@_requires_auth
def api_mark_replied(vid):
    session = _session()
    v = session.get(Vacancy, vid)
    if not v:
        session.close()
        return jsonify({"error": "not found"}), 404
    v.status = VacancyStatus.replied
    app_rec = (
        session.query(Application)
        .filter_by(vacancy_id=vid)
        .order_by(desc(Application.sent_at))
        .first()
    )
    if app_rec:
        app_rec.replied_at = datetime.now(timezone.utc)
    session.commit()
    session.close()
    return jsonify({"ok": True})


@app.route("/api/hide/<int:vid>", methods=["POST"])
@_requires_auth
def api_hide(vid):
    session = _session()
    v = session.get(Vacancy, vid)
    if v:
        v.status = VacancyStatus.rejected
        session.commit()
    session.close()
    return jsonify({"ok": True})


@app.route("/api/restore/<int:vid>", methods=["POST"])
@_requires_auth
def api_restore(vid):
    session = _session()
    v = session.get(Vacancy, vid)
    if v:
        v.status = VacancyStatus.matched
        session.commit()
    session.close()
    return jsonify({"ok": True})


@app.route("/api/resume/<int:vid>")
@_requires_auth
def api_resume(vid):
    session = _session()
    v = session.get(Vacancy, vid)
    if not v:
        session.close()
        abort(404)
    _, best_profile = _best_score(v)
    session.close()

    profile_map = {
        "Senior AI PM": "ai_pm",
        "CPO / Head of Product": "cpo",
        "Senior PM/PO": "pm",
    }
    key = profile_map.get(best_profile, "pm")
    resume_dir = Path(__file__).parent.parent.parent / "config" / "resumes"
    pdf_path = resume_dir / f"{key}.pdf"
    if not pdf_path.exists():
        abort(404)
    return send_file(str(pdf_path), as_attachment=True,
                     download_name=f"resume_{key}.pdf")


# ── channels page ─────────────────────────────────────────────────────────────

@app.route("/channels")
@_requires_auth
def channels_page():
    return render_template("channels.html")


@app.route("/api/channels")
@_requires_auth
def api_channels():
    session = _session()
    channels = session.query(Channel).order_by(desc(Channel.added_at)).all()
    result = []
    for c in channels:
        result.append({
            "id": c.id,
            "handle": c.handle,
            "title": c.title or c.handle,
            "niche": c.niche or "—",
            "active": c.active,
            "source": c.source or "manual",
            "subscriber_count": c.subscriber_count or 0,
            "post_count_30d": c.post_count_30d or 0,
            "added_at": c.added_at.strftime("%d.%m.%Y") if c.added_at else "—",
        })
    session.close()
    return jsonify({"channels": result})


@app.route("/api/channels/add_manual", methods=["POST"])
@_requires_auth
def api_add_channel_manual():
    data = request.get_json()
    handle = (data.get("handle") or "").strip().lstrip("@")
    title = (data.get("title") or "").strip()
    niche = (data.get("niche") or "product").strip()
    if not handle:
        return jsonify({"ok": False, "error": "handle обязателен"}), 400
    from jobsignal.agents.channel_finder import ChannelFinder
    finder = ChannelFinder()
    result = finder.add_manual(handle, title=title, niche=niche)
    return jsonify(result)


@app.route("/api/channels/toggle/<int:cid>", methods=["POST"])
@_requires_auth
def api_toggle_channel(cid):
    session = _session()
    c = session.get(Channel, cid)
    if not c:
        session.close()
        return jsonify({"error": "not found"}), 404
    c.active = not c.active
    session.commit()
    session.close()
    return jsonify({"ok": True, "active": c.active})


@app.route("/api/channels/delete/<int:cid>", methods=["DELETE"])
@_requires_auth
def api_delete_channel(cid):
    session = _session()
    c = session.get(Channel, cid)
    if c:
        session.delete(c)
        session.commit()
    session.close()
    return jsonify({"ok": True})


# ── candidates ────────────────────────────────────────────────────────────────

@app.route("/api/candidates")
@_requires_auth
def api_candidates():
    session = _session()
    cands = (
        session.query(ChannelCandidate)
        .filter_by(status="pending")
        .order_by(desc(ChannelCandidate.subscriber_count))
        .all()
    )
    result = [
        {
            "id": c.id,
            "handle": c.handle,
            "title": c.title or c.handle,
            "description": (c.description or "")[:120],
            "subscriber_count": c.subscriber_count or 0,
            "source": c.source,
            "search_query": c.search_query or "",
            "found_at": c.found_at.strftime("%d.%m.%Y") if c.found_at else "—",
        }
        for c in cands
    ]
    session.close()
    return jsonify({"candidates": result})


@app.route("/api/candidates/add/<int:cid>", methods=["POST"])
@_requires_auth
def api_add_candidate(cid):
    data = request.get_json() or {}
    niche = data.get("niche", "product")
    from jobsignal.agents.channel_finder import ChannelFinder
    finder = ChannelFinder()
    ok = finder.add_channel_from_candidate(cid, niche=niche)
    return jsonify({"ok": ok})


@app.route("/api/candidates/reject/<int:cid>", methods=["POST"])
@_requires_auth
def api_reject_candidate(cid):
    from jobsignal.agents.channel_finder import ChannelFinder
    finder = ChannelFinder()
    ok = finder.reject_candidate(cid)
    return jsonify({"ok": ok})


@app.route("/api/search_channels", methods=["POST"])
@_requires_auth
def api_search_channels():
    """Trigger channel search with given queries."""
    data = request.get_json() or {}
    queries = data.get("queries", [])
    if not queries:
        return jsonify({"error": "queries обязателен"}), 400
    from jobsignal.agents.channel_finder import ChannelFinder
    finder = ChannelFinder(queries=queries)
    result = finder.run(queries=queries)
    return jsonify(result)


# ── resumes page ─────────────────────────────────────────────────────────────

RESUME_DIR = Path(__file__).parent.parent.parent / "config" / "resumes"
PROFILE_KEYS = {"ai_pm", "cpo", "pm"}


@app.route("/resumes")
@_requires_auth
def resumes_page():
    return render_template("resumes.html")


@app.route("/api/resumes/status")
@_requires_auth
def api_resumes_status():
    RESUME_DIR.mkdir(parents=True, exist_ok=True)
    result = {}
    for key in PROFILE_KEYS:
        path = RESUME_DIR / f"{key}.pdf"
        if path.exists():
            stat = path.stat()
            result[key] = {
                "exists": True,
                "filename": f"{key}.pdf",
                "size_kb": round(stat.st_size / 1024),
                "updated": datetime.fromtimestamp(stat.st_mtime).strftime("%d.%m.%Y %H:%M"),
            }
        else:
            result[key] = {"exists": False}
    return jsonify({"resumes": result})


@app.route("/api/resumes/upload", methods=["POST"])
@_requires_auth
def api_resumes_upload():
    profile_key = request.form.get("profile_key", "").strip()
    if profile_key not in PROFILE_KEYS:
        return jsonify({"ok": False, "error": "Неверный profile_key"}), 400

    file = request.files.get("file")
    if not file or not file.filename:
        return jsonify({"ok": False, "error": "Файл не выбран"}), 400
    if not file.filename.lower().endswith(".pdf"):
        return jsonify({"ok": False, "error": "Только PDF"}), 400

    RESUME_DIR.mkdir(parents=True, exist_ok=True)
    dest = RESUME_DIR / f"{profile_key}.pdf"

    # backup previous
    if dest.exists():
        backup = RESUME_DIR / f"{profile_key}_prev.pdf"
        dest.rename(backup)

    try:
        file.save(str(dest))
        size_kb = round(dest.stat().st_size / 1024)
        log.info("resume uploaded: %s (%d КБ)", dest.name, size_kb)
        # auto-parse resume text and save to DB
        try:
            from jobsignal.agents.resume_parser import parse_and_save
            parse_result = parse_and_save(profile_key)
            log.info("[resumes] parsed %s: %s chars", profile_key, parse_result.get("chars", 0))
        except Exception as pe:
            log.warning("[resumes] parse error: %s", pe)
        return jsonify({"ok": True, "size_kb": size_kb})
    except Exception as exc:
        log.error("resume upload error: %s", exc)
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/api/resumes/download/<key>")
@_requires_auth
def api_resumes_download(key):
    if key not in PROFILE_KEYS:
        abort(404)
    path = RESUME_DIR / f"{key}.pdf"
    if not path.exists():
        abort(404)
    return send_file(str(path), as_attachment=True, download_name=f"resume_{key}.pdf")


# ── pipeline trigger ──────────────────────────────────────────────────────────

@app.route("/api/pipeline/run", methods=["POST"])
@_requires_auth
def api_pipeline_run():
    """Trigger collect → parse → dedup → match in background."""
    import threading
    import subprocess
    import sys

    def _run():
        python = sys.executable
        base = str(Path(__file__).parent.parent.parent)
        for cmd in ["collect", "parse", "dedup", "match"]:
            try:
                subprocess.run(
                    [python, "run.py", cmd],
                    cwd=base, timeout=600, capture_output=True
                )
            except Exception as exc:
                log.error("pipeline step %s error: %s", cmd, exc)

    t = threading.Thread(target=_run, daemon=True)
    t.start()
    return jsonify({"ok": True, "message": "Конвейер запущен в фоне (~3-5 мин)"})




# ── recruiters ────────────────────────────────────────────────────────────────

@app.route("/recruiters")
@_requires_auth
def recruiters_page():
    return render_template("recruiters.html")


@app.route("/api/recruiters")
@_requires_auth
def api_recruiters():
    import sqlite3
    try:
        conn = sqlite3.connect("data/jobsignal.db")
        rows = conn.execute("""
            SELECT handle, name, company, source, tg_handle,
                   linkedin_url, email, vacancy_count, last_seen_at, notes
            FROM recruiters
            ORDER BY vacancy_count DESC, last_seen_at DESC
            LIMIT 500
        """).fetchall()
        conn.close()
    except Exception:
        rows = []
    result = [
        {"handle": r[0], "name": r[1], "company": r[2], "source": r[3],
         "tg_handle": r[4], "linkedin_url": r[5], "email": r[6],
         "vacancy_count": r[7], "last_seen_at": r[8], "notes": r[9]}
        for r in rows
    ]
    return jsonify({"recruiters": result})


@app.route("/api/recruiters/build", methods=["POST"])
@_requires_auth
def api_recruiters_build():
    from jobsignal.agents.recruiter_builder import RecruiterBuilder
    result = RecruiterBuilder().build()
    return jsonify(result)


@app.route("/api/recruiters/update", methods=["POST"])
@_requires_auth
def api_recruiters_update():
    import sqlite3
    data = request.get_json() or {}
    handle = data.get("handle")
    if not handle:
        return jsonify({"ok": False, "error": "handle обязателен"}), 400
    conn = sqlite3.connect("data/jobsignal.db")
    conn.execute("""
        UPDATE recruiters SET
            name=COALESCE(NULLIF(?,''),name),
            company=COALESCE(NULLIF(?,''),company),
            tg_handle=COALESCE(NULLIF(?,''),tg_handle),
            linkedin_url=COALESCE(NULLIF(?,''),linkedin_url),
            email=COALESCE(NULLIF(?,''),email),
            notes=COALESCE(NULLIF(?,''),notes)
        WHERE handle=?
    """, (data.get("name",""), data.get("company",""), data.get("tg_handle",""),
          data.get("linkedin_url",""), data.get("email",""), data.get("notes",""), handle))
    conn.commit()
    conn.close()
    return jsonify({"ok": True})


@app.route("/api/recruiters/for_company")
@_requires_auth
def api_recruiters_for_company():
    company = request.args.get("company", "")
    from jobsignal.agents.recruiter_builder import RecruiterBuilder
    result = RecruiterBuilder().find_for_company(company)
    return jsonify({"recruiters": result})



# ── analytics ─────────────────────────────────────────────────────────────────

@app.route("/analytics")
@_requires_auth
def analytics_page():
    return render_template("analytics.html")


@app.route("/api/analytics")
@_requires_auth
def api_analytics():
    import sqlite3
    from datetime import datetime, timezone, timedelta

    conn = sqlite3.connect("data/jobsignal.db")
    conn.row_factory = sqlite3.Row
    now = datetime.now(timezone.utc)

    # ensure extra status columns exist
    cols = [r[1] for r in conn.execute("PRAGMA table_info(vacancies)").fetchall()]
    for col in ["interview_at", "offer_at", "rejected_at"]:
        if col not in cols:
            conn.execute(f"ALTER TABLE vacancies ADD COLUMN {col} DATETIME")
    conn.commit()

    # total matched
    total_matched = conn.execute(
        "SELECT COUNT(*) FROM vacancies WHERE status IN ('matched','drafted','applied','replied','interview','offer','rejected','no_reply') AND is_primary=1"
    ).fetchone()[0]

    # funnel counts
    funnel_statuses = ["applied", "replied", "interview", "offer", "rejected"]
    funnel = []
    for st in funnel_statuses:
        count = conn.execute(
            "SELECT COUNT(*) FROM vacancies WHERE status=? AND is_primary=1", (st,)
        ).fetchone()[0]
        if count > 0 or st in ("applied", "replied"):
            funnel.append({"status": st, "count": count})

    total_applied = next((f["count"] for f in funnel if f["status"] == "applied"), 0)
    total_replied = next((f["count"] for f in funnel if f["status"] == "replied"), 0)

    # avg reply time (days)
    avg_reply_days = None
    rows = conn.execute("""
        SELECT v.id, a.sent_at, a.replied_at
        FROM applications a
        JOIN vacancies v ON v.id = a.vacancy_id
        WHERE a.sent_at IS NOT NULL AND a.replied_at IS NOT NULL
    """).fetchall()
    if rows:
        deltas = []
        for r in rows:
            try:
                sent = datetime.fromisoformat(r["sent_at"].replace("Z", "+00:00"))
                replied = datetime.fromisoformat(r["replied_at"].replace("Z", "+00:00"))
                deltas.append((replied - sent).days)
            except Exception:
                pass
        if deltas:
            avg_reply_days = round(sum(deltas) / len(deltas), 1)

    # no reply: applied 5+ days ago, no response
    no_reply_cutoff = (now - timedelta(days=5)).isoformat()
    no_reply_rows = conn.execute("""
        SELECT v.id AS vacancy_id, v.role, v.company, v.link,
               a.sent_at,
               CAST(julianday('now') - julianday(a.sent_at) AS INTEGER) AS days
        FROM vacancies v
        JOIN applications a ON a.vacancy_id = v.id
        WHERE v.status = 'applied'
          AND a.sent_at < ?
          AND a.replied_at IS NULL
        ORDER BY days DESC
        LIMIT 20
    """, (no_reply_cutoff,)).fetchall()
    no_reply = [dict(r) for r in no_reply_rows]

    # auto-mark no_reply in DB
    for r in no_reply:
        conn.execute(
            "UPDATE vacancies SET status='no_reply' WHERE id=? AND status='applied'",
            (r["vacancy_id"],)
        )
    conn.commit()

    # applications history
    apps_rows = conn.execute("""
        SELECT v.id AS vacancy_id, v.role, v.company, v.status, v.link,
               a.sent_at, a.replied_at,
               CAST(julianday('now') - julianday(COALESCE(a.sent_at, v.created_at)) AS INTEGER) AS days_since
        FROM vacancies v
        LEFT JOIN applications a ON a.vacancy_id = v.id
        WHERE v.status IN ('applied','replied','interview','offer','rejected','no_reply')
          AND v.is_primary = 1
        ORDER BY a.sent_at DESC NULLS LAST
        LIMIT 100
    """).fetchall()

    def fmt_date(d):
        if not d:
            return None
        try:
            return datetime.fromisoformat(d.replace("Z", "+00:00")).strftime("%d.%m.%Y")
        except Exception:
            return d[:10] if d else None

    applications = []
    for r in apps_rows:
        applications.append({
            "vacancy_id": r["vacancy_id"],
            "role": r["role"],
            "company": r["company"],
            "status": r["status"],
            "link": r["link"],
            "sent_at": fmt_date(r["sent_at"]),
            "replied_at": fmt_date(r["replied_at"]),
            "days_since": r["days_since"],
        })

    # top channels by applications
    ch_rows = conn.execute("""
        SELECT c.handle, COUNT(*) AS count
        FROM vacancies v
        JOIN channels c ON c.id = v.channel_id
        WHERE v.status IN ('applied','replied','interview','offer','rejected','no_reply')
        GROUP BY c.handle
        ORDER BY count DESC
        LIMIT 10
    """).fetchall()
    top_channels = [{"handle": r["handle"], "count": r["count"]} for r in ch_rows]

    # by profile
    profile_rows = conn.execute("""
        SELECT ms.profile_key AS profile, COUNT(*) AS count
        FROM vacancies v
        JOIN match_scores ms ON ms.vacancy_id = v.id
        WHERE v.status IN ('applied','replied','interview','offer','rejected','no_reply')
          AND ms.score = (SELECT MAX(ms2.score) FROM match_scores ms2 WHERE ms2.vacancy_id = v.id)
        GROUP BY ms.profile_key
        ORDER BY count DESC
    """).fetchall()
    by_profile = [{"profile": r["profile"], "count": r["count"]} for r in profile_rows]

    conn.close()
    return jsonify({
        "total_matched": total_matched,
        "total_applied": total_applied,
        "total_replied": total_replied,
        "avg_reply_days": avg_reply_days,
        "funnel": funnel,
        "no_reply": no_reply,
        "applications": applications,
        "top_channels": top_channels,
        "by_profile": by_profile,
    })


@app.route("/api/mark_status", methods=["POST"])
@_requires_auth
def api_mark_status():
    import sqlite3
    from datetime import datetime, timezone
    data = request.get_json() or {}
    vid = data.get("vacancy_id")
    status = data.get("status")
    allowed = {"replied", "interview", "offer", "rejected", "no_reply"}
    if not vid or status not in allowed:
        return jsonify({"ok": False, "error": "invalid"}), 400

    now = datetime.now(timezone.utc).isoformat()
    conn = sqlite3.connect("data/jobsignal.db")
    conn.execute("UPDATE vacancies SET status=? WHERE id=?", (status, vid))

    # update application record
    if status == "replied":
        conn.execute(
            "UPDATE applications SET replied_at=? WHERE vacancy_id=? AND replied_at IS NULL",
            (now, vid)
        )
    conn.commit()
    conn.close()
    return jsonify({"ok": True})



# ── cover templates ───────────────────────────────────────────────────────────

DEFAULT_TEMPLATES = [
    {
        "name": "Метрика-крючок",
        "description": "Для TG. Первая строка — цифра которая цепляет",
        "is_active": True,
        "use_for_tg": True,
        "use_for_hh": False,
        "system_prompt": (
            "Напиши короткое сопроводительное письмо (5-7 предложений). "
            "Используй ТОЛЬКО факты из резюме кандидата — никаких выдуманных цифр и компаний. "
            "Стиль: деловой, живой, как человек пишет человеку. Без шаблонных фраз типа "
            "\"уверен что мой опыт\", \"стремлюсь к развитию\", \"буду полезен\". "
            "Структура: приветствие → одна строка о вакансии → кто я + компании → "
            "1-2 конкретных результата из резюме → призыв обсудить. "
            "Начни с \"Добрый день\" или \"Добрый день, {name}\" если есть имя."
        ),
        "body_template": (
            "Добрый день{name_part}\n\nЗаинтересовала вакансия {role}{company_part}.\n\n{profile_line} — работал в {companies}.\n\nИз последнего: {metric_1}. Также {metric_2}.\n\nБуду рад обсудить."
        ),
    },
    {
        "name": "Под запрос вакансии",
        "description": "Для hh.ru. Зеркалю ключевые требования вакансии",
        "is_active": False,
        "use_for_tg": False,
        "use_for_hh": True,
        "system_prompt": (
            "Напиши сопроводительное письмо (5-7 предложений). "
            "Используй ТОЛЬКО факты из резюме — без выдумок. "
            "Стиль: деловой, без канцелярита и шаблонных фраз. "
            "Ключевой приём: в начале упомяни конкретное требование из вакансии "
            "и сразу покажи что у тебя это есть. "
            "Структура: приветствие → я вижу вы ищете X → у меня есть это + компании → "
            "конкретный результат → прикладываю резюме, готов обсудить."
        ),
        "body_template": (
            "Добрый день{name_part}\n\nОткликаюсь на вакансию {role}{company_part}.\n\nВижу, вы ищете {profile_line} — у меня есть релевантный опыт в {companies}.\n\n{metric_1}. {metric_2}.\n\nПрикладываю резюме, буду рад обсудить детали."
        ),
    },
    {
        "name": "Хендз-он билдер",
        "description": "Для стартапов. Акцент: делаю сам, быстро, без армии",
        "is_active": False,
        "use_for_tg": True,
        "use_for_hh": True,
        "system_prompt": (
            "Напиши сопроводительное письмо (5-7 предложений). "
            "Используй ТОЛЬКО факты из резюме — никаких выдуманных цифр. "
            "Ключевой акцент: кандидат строит продукт сам, прототипирует на AI-стеке, "
            "умеет работать в условиях ограниченных ресурсов. "
            "Стиль: энергичный, конкретный, без воды. "
            "Структура: приветствие → вакансия → кто я в 1 строку → "
            "что умею делать сам (AI-стек, данные, прототипы) → результат → готов обсудить."
        ),
        "body_template": (
            "Добрый день{name_part}\n\nИнтересна позиция {role}{company_part}.\n\n{profile_line}, строю продукт сам — {companies}.\n\n{metric_1}. Прототипирую на AI-стеке без разработчиков.\n\nГотов обсудить."
        ),
    },
]


def _ensure_cover_templates_table():
    import sqlite3
    conn = sqlite3.connect("data/jobsignal.db")
    conn.execute("""
        CREATE TABLE IF NOT EXISTS cover_templates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name VARCHAR(128) NOT NULL,
            description TEXT,
            is_active INTEGER DEFAULT 0,
            use_for_tg INTEGER DEFAULT 1,
            use_for_hh INTEGER DEFAULT 1,
            system_prompt TEXT,
            body_template TEXT,
            created_at DATETIME DEFAULT (datetime('now'))
        )
    """)
    # seed defaults if empty
    count = conn.execute("SELECT COUNT(*) FROM cover_templates").fetchone()[0]
    if count == 0:
        for t in DEFAULT_TEMPLATES:
            conn.execute("""
                INSERT INTO cover_templates
                (name, description, is_active, use_for_tg, use_for_hh, system_prompt, body_template)
                VALUES (?,?,?,?,?,?,?)
            """, (t["name"], t["description"], int(t["is_active"]),
                  int(t["use_for_tg"]), int(t["use_for_hh"]),
                  t["system_prompt"], t["body_template"]))
    conn.commit()
    conn.close()


@app.route("/cover-templates")
@_requires_auth
def cover_templates_page():
    _ensure_cover_templates_table()
    return render_template("cover_templates.html")


@app.route("/api/cover-templates")
@_requires_auth
def api_cover_templates():
    import sqlite3
    _ensure_cover_templates_table()
    conn = sqlite3.connect("data/jobsignal.db")
    conn.row_factory = sqlite3.Row
    rows = conn.execute("SELECT * FROM cover_templates ORDER BY id").fetchall()
    conn.close()
    return jsonify({"templates": [dict(r) for r in rows]})


@app.route("/api/cover-templates/update", methods=["POST"])
@_requires_auth
def api_cover_templates_update():
    import sqlite3
    data = request.get_json() or {}
    tid = data.get("id")
    if not tid:
        return jsonify({"ok": False}), 400
    conn = sqlite3.connect("data/jobsignal.db")
    conn.execute(
        "UPDATE cover_templates SET system_prompt=?, body_template=? WHERE id=?",
        (data.get("system_prompt", ""), data.get("body_template", ""), tid)
    )
    conn.commit()
    conn.close()
    return jsonify({"ok": True})


@app.route("/api/cover-templates/set-active", methods=["POST"])
@_requires_auth
def api_cover_templates_set_active():
    import sqlite3
    data = request.get_json() or {}
    tid = data.get("id")
    if not tid:
        return jsonify({"ok": False}), 400
    conn = sqlite3.connect("data/jobsignal.db")
    conn.execute("UPDATE cover_templates SET is_active=0")
    conn.execute("UPDATE cover_templates SET is_active=1 WHERE id=?", (tid,))
    conn.commit()
    conn.close()
    return jsonify({"ok": True})


@app.route("/api/cover-templates/active")
@_requires_auth
def api_cover_templates_active():
    """Get active template — used by notify_bot for cover generation."""
    import sqlite3
    _ensure_cover_templates_table()
    conn = sqlite3.connect("data/jobsignal.db")
    conn.row_factory = sqlite3.Row
    row = conn.execute(
        "SELECT * FROM cover_templates WHERE is_active=1 LIMIT 1"
    ).fetchone()
    conn.close()
    if not row:
        return jsonify({"template": None})
    return jsonify({"template": dict(row)})

# ── serve ─────────────────────────────────────────────────────────────────────

def serve(host: str = "0.0.0.0", port: int = 5000):
    try:
        from waitress import serve as waitress_serve
        log.info("jobsignal dashboard on %s:%s", host, port)
        waitress_serve(app, host=host, port=port)
    except ImportError:
        app.run(host=host, port=port)
