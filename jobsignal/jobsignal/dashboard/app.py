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
from jobsignal.letters import ACCENTS as STYLES
from jobsignal.config import OutreachMode
from jobsignal.ratelimit import rate_status

log = logging.getLogger("jobsignal")

app = Flask(__name__)

_sf = None

def _session():
    global _sf
    if _sf is None:
        _sf = get_session_factory()
    return _sf()


# ── auth ──────────────────────────────────────────────────────────────────────

# Кабинет открыт в интернет через Caddy (deploy/expose_dashboard.sh), поэтому
# вход строже, чем во времена SSH-туннеля:
#   * без DASHBOARD_PASS пускаем только прямые запросы с самого сервера;
#     пришедшее через Caddy (есть X-Forwarded-For) без пароля не пускаем;
#   * пароль сравниваем за постоянное время;
#   * после AUTH_MAX_FAILS неверных паролей с одного адреса за AUTH_WINDOW_S
#     секунд адрес получает 429 до конца окна — перебор становится бесполезным.
AUTH_MAX_FAILS = int(os.environ.get("DASHBOARD_MAX_FAILS", "10"))
AUTH_WINDOW_S = int(os.environ.get("DASHBOARD_FAIL_WINDOW_S", "900"))
_auth_fails: dict[str, list[float]] = {}


def _client_ip() -> str:
    """Адрес клиента. X-Forwarded-For верим, только если запрос пришёл от
    Caddy с этого же сервера: иначе заголовок подделает кто угодно."""
    remote = request.remote_addr or ""
    fwd = request.headers.get("X-Forwarded-For", "")
    if remote in ("127.0.0.1", "::1") and fwd:
        return fwd.split(",")[0].strip()
    return remote


def _is_local_direct() -> bool:
    return (request.remote_addr in ("127.0.0.1", "::1")
            and not request.headers.get("X-Forwarded-For"))


def _check_auth(username: str, password: str) -> bool:
    import hmac
    u = os.environ.get("DASHBOARD_USER", "admin")
    p = os.environ.get("DASHBOARD_PASS", "")
    if not p:
        return _is_local_direct()
    return (hmac.compare_digest(username.encode(), u.encode())
            and hmac.compare_digest(password.encode(), p.encode()))


def _locked_out(ip: str) -> bool:
    import time
    now = time.time()
    fails = [t for t in _auth_fails.get(ip, []) if now - t < AUTH_WINDOW_S]
    _auth_fails[ip] = fails
    return len(fails) >= AUTH_MAX_FAILS


def _note_fail(ip: str) -> None:
    import time
    _auth_fails.setdefault(ip, []).append(time.time())
    if len(_auth_fails[ip]) == AUTH_MAX_FAILS:
        log.warning("[dashboard] %d неверных входов с %s — адрес закрыт на %d мин",
                    AUTH_MAX_FAILS, ip, AUTH_WINDOW_S // 60)


def _requires_auth(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        ip = _client_ip()
        if _locked_out(ip):
            return Response("Слишком много неверных паролей, попробуй позже",
                            429, {"Retry-After": str(AUTH_WINDOW_S)})
        if not os.environ.get("DASHBOARD_PASS") and not _is_local_direct():
            return Response("Кабинет закрыт: на сервере не задан DASHBOARD_PASS",
                            503)
        auth = request.authorization
        if not auth or not _check_auth(auth.username or "", auth.password or ""):
            if auth:
                _note_fail(ip)
            return Response(
                "Нужна авторизация", 401,
                {"WWW-Authenticate": 'Basic realm="JobSignal", charset="UTF-8"'}
            )
        return f(*args, **kwargs)
    return decorated


@app.after_request
def _security_headers(resp):
    resp.headers.setdefault("X-Frame-Options", "DENY")
    resp.headers.setdefault("X-Content-Type-Options", "nosniff")
    resp.headers.setdefault("Referrer-Policy", "no-referrer")
    resp.headers.setdefault("X-Robots-Tag", "noindex, nofollow")
    return resp


app.before_request(lambda: None)  # placeholder; auth applied per-route


# ── helpers ───────────────────────────────────────────────────────────────────

def _best_score(vacancy) -> tuple[int, str]:
    if not vacancy.match_scores:
        return 0, ""
    # score в схеме nullable: код всегда пишет int, но одна строка с NULL из
    # прошлого или из правки руками уронила бы сравнение в max() на TypeError,
    # то есть всю страницу дашборда. Дешевле считать пустой балл нулём.
    # Сначала действующие профили (config/profiles.yaml): оценки снятых
    # профилей остаются в базе у непереоценённых вакансий, но решать по ним,
    # каким резюме откликаться, нельзя. Нет действующих — показываем что есть.
    from jobsignal import profiles as registry
    active = set(registry.names())
    pool = [s for s in vacancy.match_scores if s.profile_key in active] \
        or list(vacancy.match_scores)
    best = max(pool, key=lambda s: s.score or 0)
    return best.score, best.profile_key


def _rate_status() -> dict:
    """Лимиты откликов. Единая реализация в jobsignal.ratelimit — её же
    использует HHApplyAgent, чтобы дашборд и автоотклик считали одну квоту."""
    return rate_status()


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
        # Ключ фильтра — имя профиля: именно оно лежит в
        # match_scores.profile_key. Раньше ключом бралось поле key, которого в
        # profiles.yaml нет: KeyError уводил в запасной словарь с ключами
        # ai_pm/cpo/pm, и фильтр по профилю всегда возвращал пустой список.
        for p in data.get("profiles", []):
            if p.get("name"):
                profiles[p["name"]] = p["name"]
    except Exception:
        profiles = {}
    if not profiles:
        from jobsignal import profiles as registry
        profiles = {n: n for n in registry.names()}
    return render_template("index.html", styles=list(STYLES.keys()),
                           style_labels={k: v["label"] for k, v in STYLES.items()},
                           profiles=profiles, outreach_mode=_outreach_mode())


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
    # Только реально отправленные: черновики (sent_at пуст) раньше тоже
    # считались откликами, и цифра в шапке была завышена.
    applied = (session.query(func.count(Application.id))
               .filter(Application.sent_at.isnot(None)).scalar())
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


def _letter_channel(v) -> str:
    """Канал письма по контакту вакансии (правила письма у каналов разные)."""
    ct = v.contact_type or detect_contact_type(v)
    return {"hh": "hh", "tg": "telegram", "form": "form", "email": "email",
            "linkedin": "email"}.get(ct, "hh")


@app.route("/api/generate_draft", methods=["POST"])
@_requires_auth
def api_generate_draft():
    """Письмо по кнопке «Сгенерировать» — тем же сервисом, что автоотклик и
    бот (jobsignal/letters.py), с выбранным акцентом и правилами канала."""
    from jobsignal import letters
    data = request.get_json() or {}
    vid = data.get("vacancy_id")
    session = _session()
    from sqlalchemy.orm import joinedload
    v = session.query(Vacancy).options(joinedload(Vacancy.match_scores)).filter_by(id=vid).first()
    if not v:
        session.close()
        return jsonify({"error": "not found"}), 404
    try:
        _, best_profile = _best_score(v)
        letter = letters.for_vacancy(v.id, channel=_letter_channel(v),
                                     accent=data.get("style"),
                                     profile=best_profile or None)
        if not letter.text:
            raise ValueError("; ".join(letter.issues) or "пустой ответ модели")
        letters.remember(v, letter)
        if v.status == VacancyStatus.matched:
            v.status = VacancyStatus.drafted
        session.commit()
        session.close()
        return jsonify({"draft": letter.text, "channel": letter.channel,
                        "profile": letter.profile, "issues": letter.issues})
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
    from jobsignal import letters
    from jobsignal import profiles as registry
    profile_name = best_profile or registry.default_name()
    channel = _letter_channel(v)
    # Письмо то же, что уходит в автоотклике и боте. Черновик берём, только
    # если он написан для этого канала и профиля по действующим правилам;
    # иначе пишем и запоминаем, чтобы повторное открытие не стоило вызова.
    cover = letters.reusable(v, channel, profile_name)
    if not cover:
        letter = letters.for_vacancy(v.id, channel=channel, profile=profile_name)
        cover = letter.text
        if letter.ok:
            letters.remember(v, letter)
            session.commit()
    data = get_outreach_data(v, profile_name, cover_text=cover or None)
    session.close()
    return jsonify(data)


@app.route("/api/mark_applied/<int:vid>", methods=["POST"])
@_requires_auth
def api_mark_applied(vid):
    session = _session()
    v = session.get(Vacancy, vid)
    if not v:
        session.close()
        return jsonify({"error": "not found"}), 404
    v.status = VacancyStatus.applied
    # Черновик, который дашборд показывал в момент нажатия. Отклик подавал
    # пользователь руками, поэтому канал — manual, а не tg: иначе ручные
    # отметки неотличимы от писем, отправленных системой, и конверсия
    # считается по выборке, в которой система ничего не отправляла.
    # Раньше сюда писалась пустая строка — 65 таких строк в базе, и по ним
    # уже не восстановить, что именно ушло рекрутёру.
    draft = (v.draft_text or "").strip()
    session.commit()
    session.close()
    try:
        import sqlite3 as _sq
        conn = _sq.connect("data/jobsignal.db")
        # Повторное нажатие не должно создавать вторую строку: индекса на
        # vacancy_id нет, и OR IGNORE от дубля здесь не спасал — а лишний
        # отклик съедал слот часовой квоты.
        already = conn.execute(
            "SELECT 1 FROM applications WHERE vacancy_id=? AND sent_at IS NOT NULL",
            (vid,)
        ).fetchone()
        if not already:
            conn.execute(
                "INSERT INTO applications (vacancy_id, message_text, channel, is_draft, sent_at, created_at) VALUES (?,?,?,?,datetime(\'now\'),datetime(\'now\'))",
                (vid, draft, "manual", 0)
            )
        conn.commit()
        conn.close()
    except Exception as _e:
        log.warning("application insert: %s", _e)
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
    warning = None
    if v:
        # Возврат в работу — единственный живой путь, который может увести
        # статус назад с applied. Строка отклика при этом остаётся: получится
        # отправленный отклик у вакансии, которая снова числится matched, и
        # воронка «отклики → ответы» перестанет сходиться. Решение за
        # пользователем, но молча это происходить не должно.
        sent = (
            session.query(Application)
            .filter(Application.vacancy_id == vid,
                    Application.sent_at.is_not(None))
            .first()
        )
        if sent is not None and v.status == VacancyStatus.applied:
            warning = ("по вакансии уже есть отправленный отклик "
                       f"от {sent.sent_at}; строка отклика сохранена")
            log.warning("restore #%s: %s", vid, warning)
        v.status = VacancyStatus.matched
        session.commit()
    session.close()
    return jsonify({"ok": True, "warning": warning})


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

    from jobsignal import profiles as registry
    prof = registry.resolve(best_profile)
    pdf_path = prof.resume_pdf() if prof else None
    if pdf_path is None:
        abort(404)
    return send_file(str(pdf_path), as_attachment=True,
                     download_name=f"resume_{prof.key}.pdf")


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
def _profile_keys() -> set[str]:
    from jobsignal import profiles as registry
    return set(registry.keys())


@app.route("/resumes")
@_requires_auth
def resumes_page():
    return render_template("resumes.html")


@app.route("/api/resumes/status")
@_requires_auth
def api_resumes_status():
    RESUME_DIR.mkdir(parents=True, exist_ok=True)
    from jobsignal.agents.resume_parser import pdf_path_for
    from jobsignal import profiles as registry
    result = {}
    meta = []
    for p in registry.all_profiles():
        meta.append({"key": p.key, "name": p.name, "label": p.letter_label,
                     "hh_resume_title": p.hh_resume_title})
    for key in _profile_keys():
        # Не RESUME_DIR/<key>.pdf: отдельных PDF по профилям больше нет, все
        # три ссылаются на мастер-резюме. Иначе страница показывала бы
        # «файла нет» там, где письма и отправка резюме работают.
        path = pdf_path_for(key)
        if path is not None:
            stat = path.stat()
            result[key] = {
                "exists": True,
                "filename": path.name,
                "size_kb": round(stat.st_size / 1024),
                "updated": datetime.fromtimestamp(stat.st_mtime).strftime("%d.%m.%Y %H:%M"),
            }
        else:
            result[key] = {"exists": False}
    return jsonify({"resumes": result, "profiles": meta})


@app.route("/api/resumes/upload", methods=["POST"])
@_requires_auth
def api_resumes_upload():
    profile_key = request.form.get("profile_key", "").strip()
    if profile_key not in _profile_keys():
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
    if key not in _profile_keys():
        abort(404)
    from jobsignal.agents.resume_parser import pdf_path_for
    path = pdf_path_for(key)
    if path is None:
        abort(404)
    return send_file(str(path), as_attachment=True, download_name=f"resume_{key}.pdf")


# ── профили: загрузка резюме, правка, удаление ───────────────────────────────
# Страница «Резюме» управляет профилями целиком: PDF заводит профиль (модель
# сама называет его и подбирает запросы поиска), профиль можно поправить,
# сделать основным и удалить. Пишется config/profiles.yaml (jobsignal/profiles.py).

def _pdf_info(path) -> dict:
    if path is None or not path.exists():
        return {"exists": False}
    st = path.stat()
    return {"exists": True, "filename": path.name, "size_kb": round(st.st_size / 1024),
            "updated": datetime.fromtimestamp(st.st_mtime).strftime("%d.%m.%Y %H:%M")}


def _profile_json(p, i: int) -> dict:
    return {"key": p.key, "name": p.name, "label": p.label,
            "hh_resume_title": p.hh_resume_title, "focus": p.focus,
            "searches": list(p.searches), "is_default": i == 0,
            "pdf": _pdf_info(p.resume_pdf())}


@app.route("/api/profiles")
@_requires_auth
def api_profiles():
    from jobsignal import profiles as registry
    return jsonify({"profiles": [_profile_json(p, i)
                                 for i, p in enumerate(registry.all_profiles())]})


@app.route("/api/profiles/upload", methods=["POST"])
@_requires_auth
def api_profiles_upload():
    """PDF резюме: новый профиль (key пуст) или замена резюме у профиля key.

    Новый профиль модель называет сама по тексту резюме и подбирает ему
    запросы поиска. rederive=1 при замене — заново определить поля профиля
    по новому резюме.
    """
    from jobsignal import profiles as registry
    from jobsignal.agents.resume_parser import extract_text, parse_and_save
    from jobsignal.profile_builder import derive

    file = request.files.get("file")
    if not file or not file.filename:
        return jsonify({"ok": False, "error": "Файл не выбран"}), 400
    if not file.filename.lower().endswith(".pdf"):
        return jsonify({"ok": False, "error": "Нужен PDF"}), 400
    key = (request.form.get("key") or "").strip()
    rederive = request.form.get("rederive") == "1"

    RESUME_DIR.mkdir(parents=True, exist_ok=True)
    incoming = RESUME_DIR / "_incoming.pdf"
    file.save(str(incoming))
    try:
        text = extract_text(str(incoming)).strip()
        if len(text) < 300:
            return jsonify({"ok": False, "error": "Из PDF не извлёкся текст — "
                            "это скан или защищённый файл. Нужен PDF с текстом."}), 400
        existing = registry.by_key(key) if key else None
        if key and existing is None:
            return jsonify({"ok": False, "error": f"Профиля {key} нет"}), 404
        fields = None
        if existing is None or rederive:
            try:
                fields = derive(text)
            except Exception as exc:  # noqa: BLE001
                return jsonify({"ok": False, "error": f"Не удалось разобрать резюме: {exc}"}), 502
        if existing is None:
            # Профиль с тем же именем уже есть, но без PDF (так бывает после
            # переноса profiles.yaml на сервер) — резюме его и ждёт: прикрепляем,
            # а не заводим дубль. Поля, заданные человеком, не трогаем.
            same = registry.by_name(fields["name"])
            if same is not None and same.resume_pdf() is None:
                existing, key, fields = same, same.key, None
        if existing is None:
            taken = set(registry.keys())
            names = set(registry.names())
            if fields["name"] in names:
                fields["name"] = f"{fields['name']} ({len(names) + 1})"
            key = registry.slug(fields["name"], taken)
            dest = RESUME_DIR / f"{key}.pdf"
            incoming.replace(dest)
            registry.add({"key": key, **fields,
                          "resume_path": f"./config/resumes/{key}.pdf"})
            action = "created"
        else:
            dest = RESUME_DIR / f"{key}.pdf"
            if dest.exists():
                dest.replace(RESUME_DIR / f"{key}_prev.pdf")
            incoming.replace(dest)
            if fields:
                _apply_profile_update(key, fields)
            action = "replaced"
        parse_and_save(key)
        idx = registry.keys().index(key)
        log.info("[profiles] %s: %s (%s)", action, key, registry.by_key(key).name)
        return jsonify({"ok": True, "action": action,
                        "profile": _profile_json(registry.by_key(key), idx)})
    finally:
        if incoming.exists():
            incoming.unlink()


def _apply_profile_update(key: str, fields: dict):
    """Поправить профиль; при смене имени перенести на новое имя оценки и
    черновики, иначе вакансии выпали бы из очереди как оценённые под
    «снятый» профиль."""
    from jobsignal import profiles as registry
    prof, old_name = registry.update(key, fields)
    if old_name and old_name != prof.name:
        session = _session()
        try:
            from sqlalchemy import text as _t
            session.execute(_t("UPDATE match_scores SET profile_key=:new WHERE profile_key=:old"),
                            {"new": prof.name, "old": old_name})
            session.execute(_t("UPDATE vacancies SET draft_profile=:new WHERE draft_profile=:old"),
                            {"new": prof.name, "old": old_name})
            session.commit()
        finally:
            session.close()
        log.info("[profiles] %s переименован: «%s» → «%s»", key, old_name, prof.name)
    return prof


@app.route("/api/profiles/<key>/update", methods=["POST"])
@_requires_auth
def api_profiles_update(key):
    from jobsignal import profiles as registry
    data = request.get_json() or {}
    fields = {}
    for f in ("name", "label", "hh_resume_title", "focus"):
        if f in data:
            fields[f] = str(data[f] or "").strip()
    if "searches" in data:
        raw = data["searches"]
        items = raw if isinstance(raw, list) else str(raw).replace(",", "\n").splitlines()
        fields["searches"] = [x.strip() for x in items if x and x.strip()][:15]
    if data.get("make_default"):
        fields["make_default"] = True
    if "name" in fields and not fields["name"]:
        return jsonify({"ok": False, "error": "Имя профиля не может быть пустым"}), 400
    if registry.by_key(key) is None:
        return jsonify({"ok": False, "error": f"Профиля {key} нет"}), 404
    try:
        prof = _apply_profile_update(key, fields)
    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400
    return jsonify({"ok": True, "profile": _profile_json(prof, registry.keys().index(key))})


@app.route("/api/profiles/<key>/delete", methods=["POST"])
@_requires_auth
def api_profiles_delete(key):
    """Удалить профиль. PDF уходит в config/resumes/archive/, оценки в базе
    остаются, но в очередь откликов больше не попадают."""
    from jobsignal import profiles as registry
    if registry.by_key(key) is None:
        return jsonify({"ok": False, "error": f"Профиля {key} нет"}), 404
    if len(registry.all_profiles()) <= 1:
        return jsonify({"ok": False, "error": "Это последний профиль — без профилей "
                        "оценивать вакансии не по чему. Сначала добавь другой."}), 400
    gone = registry.remove(key)
    pdf = RESUME_DIR / f"{key}.pdf"
    if pdf.exists():
        arch = RESUME_DIR / "archive"
        arch.mkdir(parents=True, exist_ok=True)
        pdf.replace(arch / f"{key}-{datetime.now():%Y%m%d-%H%M}.pdf")
    log.info("[profiles] удалён %s (%s)", key, gone.name)
    return jsonify({"ok": True})


def _rescore_candidates(days: int = 14) -> int:
    """Сколько свежих вакансий hh не оценены под все действующие профили."""
    from jobsignal import profiles as registry
    names = registry.names()
    if not names:
        return 0
    session = _session()
    try:
        from sqlalchemy import text as _t
        marks = ",".join(f":n{i}" for i in range(len(names)))
        params = {f"n{i}": n for i, n in enumerate(names)}
        params["days"] = f"-{days} days"
        params["cnt"] = len(names)
        return session.execute(_t(f"""
            select count(*) from vacancies v
             where v.is_primary = 1 and v.contact_type = 'hh'
               and v.created_at > datetime('now', :days)
               and exists (select 1 from match_scores m where m.vacancy_id = v.id)
               and (select count(distinct m.profile_key) from match_scores m
                     where m.vacancy_id = v.id and m.profile_key in ({marks})) < :cnt
               and not exists (select 1 from applications a
                                where a.vacancy_id = v.id and a.sent_at is not null)
        """), params).scalar() or 0
    finally:
        session.close()


@app.route("/api/profiles/rescore", methods=["GET", "POST"])
@_requires_auth
def api_profiles_rescore():
    """GET — сколько вакансий за 14 дней надо переоценить и почём; POST —
    запустить переоценку в фоне (tools/rescore.py --profiles-changed)."""
    import fcntl
    import subprocess
    import sys
    import threading

    n = _rescore_candidates()
    if request.method == "GET":
        return jsonify({"count": n, "cost_usd": round(n * 0.006, 2)})
    from jobsignal import profiles as registry
    from jobsignal.config import get_config
    if len(get_config().profiles) < len(registry.all_profiles()):
        return jsonify({"ok": False, "message": "Не у всех профилей загружено резюме — "
                        "переоценивать рано"})
    if not n:
        return jsonify({"ok": True, "message": "Переоценивать нечего: все свежие вакансии "
                        "оценены под действующие профили"})
    lock_path = os.environ.get("HH_LOCK_PATH", "/run/jobsignal-hh.lock")
    try:
        fh = open(lock_path, "w")
        fcntl.flock(fh, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except BlockingIOError:
        fh.close()
        return jsonify({"ok": False, "message": "Сейчас идёт прогон — попробуй через "
                        "несколько минут"})

    def _run():
        base = str(Path(__file__).parent.parent.parent)
        try:
            subprocess.run([sys.executable, "tools/rescore.py", "--profiles-changed",
                            "--days", "14", "--limit", str(min(n, 600)), "--apply"],
                           cwd=base, timeout=3600, capture_output=True)
        except Exception as exc:  # noqa: BLE001
            log.error("rescore error: %s", exc)
        finally:
            fcntl.flock(fh, fcntl.LOCK_UN)
            fh.close()

    threading.Thread(target=_run, daemon=True).start()
    return jsonify({"ok": True, "message": f"Переоценка {min(n, 600)} вакансий запущена в "
                    f"фоне, займёт 10-30 минут"})


# ── pipeline trigger ──────────────────────────────────────────────────────────

@app.route("/api/pipeline/run", methods=["POST"])
@_requires_auth
def api_pipeline_run():
    """Trigger collect → parse → dedup → match in background."""
    import threading
    import subprocess
    import sys

    import fcntl

    # Тот же замок, что у часового прогона и автоотклика. Без него кнопка
    # запускала разбор и оценку параллельно автоотклику, оценка держала
    # запись в базе, и запись отклика падала на «database is locked» (16.10).
    lock_path = os.environ.get("HH_LOCK_PATH", "/run/jobsignal-hh.lock")
    try:
        fh = open(lock_path, "w")
        fcntl.flock(fh, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except BlockingIOError:
        fh.close()
        return jsonify({"ok": False, "message": "Прогон уже идёт (часовой "
                        "конвейер или автоотклик) — данные обновятся сами"})
    except OSError as exc:
        return jsonify({"ok": False, "message": f"Замок прогона недоступен: {exc}"})

    def _run():
        python = sys.executable
        base = str(Path(__file__).parent.parent.parent)
        try:
            for cmd in ["collect", "parse", "dedup", "match"]:
                try:
                    subprocess.run(
                        [python, "run.py", cmd],
                        cwd=base, timeout=600, capture_output=True
                    )
                except Exception as exc:
                    log.error("pipeline step %s error: %s", cmd, exc)
        finally:
            fcntl.flock(fh, fcntl.LOCK_UN)
            fh.close()

    t = threading.Thread(target=_run, daemon=True)
    t.start()
    return jsonify({"ok": True, "message": "Конвейер запущен в фоне (~3-5 мин)"})




# ── hh.ru автоотклик ──────────────────────────────────────────────────────────

def _outreach_mode() -> str:
    """semi_auto (дефолт) | full_auto. Читается из OUTREACH_MODE в config/.env."""
    raw = (os.environ.get("OUTREACH_MODE") or "").split("#")[0].strip().lower()
    return raw if raw in (OutreachMode.SEMI_AUTO.value, OutreachMode.FULL_AUTO.value) \
        else OutreachMode.SEMI_AUTO.value


# Состояние фонового прогона: один в момент времени.
_HH_JOB: dict = {"running": False, "started_at": None, "mode": None,
                 "limit": None, "result": None, "error": None}


@app.route("/api/hh_apply/status")
@_requires_auth
def api_hh_apply_status():
    """Состояние прогона + очередь hh.ru + последний отчёт."""
    session = _session()
    queue = session.query(func.count(Vacancy.id)).filter(
        Vacancy.is_primary == True,
        Vacancy.contact_type == "hh",
        Vacancy.status.in_([VacancyStatus.matched, VacancyStatus.drafted]),
    ).scalar() or 0
    session.close()

    last = _HH_JOB.get("result")
    if last is None:
        try:
            last = json.loads(Path("data/hh_apply_last.json").read_text(encoding="utf-8"))
        except Exception:
            last = None
    return jsonify({
        "running": _HH_JOB["running"],
        "mode": _outreach_mode(),
        "queue": queue,
        "rate": _rate_status(),
        "job": {k: _HH_JOB[k] for k in ("started_at", "mode", "limit", "error")},
        "last": last,
    })


@app.route("/api/hh_apply/run", methods=["POST"])
@_requires_auth
def api_hh_apply_run():
    """Запуск автоотклика на hh.ru в фоне.

    body: {send: bool, limit: int, confirm: bool}
      send=false (дефолт) — dry-run: hh.ru открывается, ничего не отправляется.
      send=true           — боевой режим, требует confirm=true при любом
                            OUTREACH_MODE.
    Подтверждение спрашивается и в full_auto. Клик по кнопке — не то же
    намерение, что автоматическая отправка по правилам: у таймера есть свой
    порог (HH_AUTO_THRESHOLD) и свой режим-выключатель, а кнопка отправляет
    по общему порогу и прямо сейчас. Режим разрешает работать таймеру, а не
    отменяет вопрос человеку.
    Число отправок всё равно ограничено общей квотой OUTREACH_PER_HOUR/DAY.
    """
    import threading

    data = request.get_json(silent=True) or {}
    send = bool(data.get("send"))
    confirm = bool(data.get("confirm"))
    mode = _outreach_mode()

    try:
        limit = int(data.get("limit") or 0) or None
    except (TypeError, ValueError):
        return jsonify({"ok": False, "error": "limit должен быть числом"}), 400

    if _HH_JOB["running"]:
        return jsonify({"ok": False, "error": "Прогон уже идёт"}), 409

    if send and not confirm:
        return jsonify({
            "ok": False, "error": "confirm_required", "mode": mode,
            "message": "Отправка по кнопке — только с подтверждением",
        }), 409

    if send:
        rate = _rate_status()
        if rate["allowed_now"] <= 0:
            return jsonify({
                "ok": False, "error": "rate_limit",
                "next_slot": rate["next_slot"], "rate": rate,
            }), 429

    from jobsignal.agents.hh_apply import HHApplyAgent
    from jobsignal.config import get_config

    _HH_JOB.update(running=True, error=None, result=None, limit=limit,
                   mode="send" if send else "dry_run",
                   started_at=datetime.now(timezone.utc).isoformat())

    def _run():
        try:
            _HH_JOB["result"] = HHApplyAgent(
                get_config(), dry_run=not send, limit=limit
            ).run()
        except Exception as exc:  # noqa: BLE001
            log.exception("hh_apply run failed")
            _HH_JOB["error"] = f"{type(exc).__name__}: {exc}"
        finally:
            _HH_JOB["running"] = False

    threading.Thread(target=_run, daemon=True).start()
    return jsonify({
        "ok": True, "mode": _HH_JOB["mode"], "limit": limit,
        "message": ("Отправляю отклики на hh.ru…" if send
                    else "Проверяю очередь на hh.ru без отправки…"),
    })


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



# ── состояние системы ─────────────────────────────────────────────────────────

@app.route("/status")
@_requires_auth
def status_page():
    return render_template("status.html")


@app.route("/api/status")
@_requires_auth
def api_status():
    from jobsignal.dashboard.status import collect
    return jsonify(collect())


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
            except Exception as exc:
                # Строка выпадает из среднего времени ответа — пусть это будет
                # видно в логе, иначе метрика тихо считается по части выборки.
                log.warning("avg_reply: вакансия #%s, даты не разобраны: %s",
                            r["id"], exc)
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

def serve(host: str | None = None, port: int = 5000):
    # По умолчанию только localhost: смотрим через SSH-туннель
    # (ssh -L 5000:localhost:5000). Раньше слушал 0.0.0.0, и при пустом
    # DASHBOARD_PASS и выключенном фаерволе база с откликами и письмами была
    # бы открыта всему интернету. Снаружи — только явным DASHBOARD_HOST.
    host = host or os.environ.get("DASHBOARD_HOST", "127.0.0.1")
    try:
        from waitress import serve as waitress_serve
        log.info("jobsignal dashboard on %s:%s", host, port)
        waitress_serve(app, host=host, port=port)
    except ImportError:
        app.run(host=host, port=port)
