"""Служебный раздел кабинета: обновление кода и проверки без SSH.

Зачем. До 08.11 каждое изменение и каждая проверка шли так: правка в
репозитории → человек копирует команды в терминал сервера → пересылает вывод
обратно. Здесь то же самое делается запросами к кабинету (по HTTPS, под тем
же паролем):

  GET  /api/admin/info               версия кода, службы, последнее обновление
  POST /api/admin/update             подтянуть ветку и перезапустить службы
                                     (deploy/self_update.sh, отдельной службой)
  POST /api/admin/run                {"cmd": "...", "args": [...]} — команда из
                                     белого списка ниже, фоном; ответ — id задачи
  GET  /api/admin/jobs/<id>          вывод и состояние задачи
  GET  /api/admin/logs?unit=…&n=200  журнал службы из белого списка

Произвольных команд нет: только перечисленные, аргументы проверяются. Боевой
отклик (hh-apply --send) сюда намеренно не входит — отклики шлёт только
таймер в пределах квоты.
"""
from __future__ import annotations

import json
import os
import re
import shlex
import subprocess
import time
import uuid
from pathlib import Path

from flask import abort, jsonify, request

APP_DIR = Path(__file__).resolve().parent.parent.parent
PY = str(APP_DIR / ".venv" / "bin" / "python")
JOBS = APP_DIR / "data" / "admin_jobs"
UPDATE_STATE = APP_DIR / "data" / "update_state.json"
UPDATE_LOG = APP_DIR / "data" / "update.log"

UNITS = ("jobsignal-bot", "jobsignal-dashboard", "jobsignal-pipeline",
         "jobsignal-hh-apply", "jobsignal-replies", "jobsignal-tgwatch",
         "jobsignal-tgscout", "jobsignal-backup", "jobsignal-update")
TIMERS = ("jobsignal-pipeline", "jobsignal-hh-apply", "jobsignal-replies",
          "jobsignal-tgwatch", "jobsignal-tgscout", "jobsignal-backup",
          "jobsignal-update")

MAX_JOBS = 2
_INT = re.compile(r"^\d{1,3}$")
_HH_VAC = re.compile(r"^https://hh\.ru/vacancy/\d{4,12}$")
_PROVIDER = re.compile(r"^(anthropic|openai|deepseek|gigachat|ollama)$")


def _build(cmd: str, args: list[str]) -> list[str]:
    """argv команды из белого списка или ValueError."""
    args = [str(a) for a in args]

    def flags(allowed: set[str]) -> list[str]:
        bad = [a for a in args if a not in allowed]
        if bad:
            raise ValueError(f"недопустимые аргументы: {bad}")
        return args

    def one_int(default: str, top: int) -> str:
        n = args[0] if args else default
        if not _INT.match(n) or not 0 < int(n) <= top:
            raise ValueError(f"нужно число от 1 до {top}")
        return n

    if cmd == "llm-check":
        return [PY, "run.py", "llm-check"]
    if cmd == "status":
        return [PY, "run.py", "status"]
    if cmd == "proxy-traffic":
        return [PY, "run.py", "proxy-traffic"]
    if cmd == "hh-session":
        return [PY, "run.py", "hh-session", *flags({"--probe"})]
    if cmd == "hh-resumes":
        return [PY, "run.py", "hh-resumes", *flags({"--refresh", "--unblock"})]
    if cmd == "hh-preview":
        return [PY, "run.py", "hh-preview", "--limit", one_int("10", 50)]
    if cmd == "hh-apply-dry":          # открывает вакансии, ничего не отправляет
        return [PY, "run.py", "hh-apply", "--limit", one_int("3", 10)]
    if cmd == "letter-preview":
        mode = args[0] if args else "new"
        if mode not in ("new", "sent"):
            raise ValueError("режим new или sent")
        args = args[1:]
        return [PY, "tools/letter_preview.py", mode, one_int("3", 10)]
    if cmd == "try-resume-choice":
        if len(args) != 1 or not _HH_VAC.match(args[0]):
            raise ValueError("нужен адрес https://hh.ru/vacancy/<номер>")
        return [PY, "tools/try_resume_choice.py", args[0]]
    if cmd == "rescore-dry":
        return [PY, "tools/rescore.py", "--profiles-changed"]
    if cmd == "llm-use":
        if not args or not _PROVIDER.match(args[0]):
            raise ValueError("нужен провайдер")
        rest = args[1:]
        if any(a != "--drop-openai" for a in rest):
            raise ValueError("допустим только --drop-openai")
        return [PY, "run.py", "llm-use", args[0], *rest]
    if cmd == "probe-sources":
        q = args[0] if args else "product owner"
        if not re.fullmatch(r"[\w .+#-]{2,60}", q):
            raise ValueError("запрос: буквы, цифры, пробелы, до 60 знаков")
        return [PY, "tools/probe_sources.py", q]
    if cmd == "probe-board":
        if not args or args[0] not in ("habr", "getmatch", "geekjob", "avito"):
            raise ValueError("площадка: habr | getmatch | geekjob | avito")
        return [PY, "tools/probe_board.py", args[0]]
    if cmd == "set-option":
        if len(args) != 2:
            raise ValueError("нужно: KEY VALUE")
        return [PY, "run.py", "set-option", args[0], args[1]]
    if cmd == "boards-collect":
        return [PY, "run.py", "boards-collect", *flags({"--force"})]
    if cmd == "profile-text":
        # Текст резюме профиля без контактов — готовить ответы рекрутёрам.
        if len(args) != 1 or not re.fullmatch(r"[a-z0-9_-]{1,32}", args[0]):
            raise ValueError("нужен ключ профиля (po, loyalty)")
        return [PY, "-c", "import sys; sys.path.insert(0, '.');"
                "from dotenv import load_dotenv; load_dotenv('config/.env');"
                "from jobsignal import profiles; from jobsignal.agents.matcher import strip_contacts;"
                f"p = profiles.by_key('{args[0]}'); print(strip_contacts(p.cv_text()) if p else 'нет профиля')"]
    if cmd == "funnel-report":
        return [PY, "tools/funnel_report.py", "--days", one_int("90", 365)]
    if cmd == "tg-scout-dry":
        return [PY, "run.py", "tg-scout", "--dry-run"]
    if cmd == "pipeline":
        return ["systemctl", "start", "--no-block", "jobsignal-pipeline.service"]
    raise ValueError(f"неизвестная команда: {cmd}")


COMMANDS = ("llm-check", "status", "proxy-traffic", "hh-session", "hh-resumes",
            "hh-preview", "hh-apply-dry", "letter-preview", "try-resume-choice",
            "rescore-dry", "llm-use", "tg-scout-dry", "pipeline", "funnel-report", "probe-sources", "probe-board", "set-option", "boards-collect", "profile-text")


def _sh(argv: list[str], timeout: int = 20) -> str:
    try:
        out = subprocess.run(argv, capture_output=True, text=True, timeout=timeout)
        return (out.stdout + out.stderr).strip()
    except Exception as exc:  # noqa: BLE001
        return f"{type(exc).__name__}: {exc}"


def _git_head() -> dict:
    src = Path("/root/visa")
    return {"commit": _sh(["git", "-C", str(src), "log", "-1",
                           "--format=%h %cd %s", "--date=format:%d.%m %H:%M"])}


def register(app, requires_auth, same_origin) -> None:

    def guard_write():
        if not same_origin():
            abort(403)

    @app.route("/api/admin/info")
    @requires_auth
    def admin_info():
        try:
            state = json.loads(UPDATE_STATE.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            state = None
        active = _sh(["systemctl", "is-active", *[u + ".service" for u in UNITS]])
        units = dict(zip(UNITS, active.splitlines()))
        timers = _sh(["systemctl", "list-timers", *[t + ".timer" for t in TIMERS],
                      "--no-pager", "--all"])
        return jsonify({"code": _git_head(), "update": state, "units": units,
                        "timers": timers, "commands": COMMANDS,
                        "disk": _sh(["df", "-h", "/"])})

    @app.route("/api/admin/update", methods=["POST"])
    @requires_auth
    def admin_update():
        guard_write()
        force = request.args.get("force") == "1"
        unit = "jobsignal-update-force.service" if force else "jobsignal-update.service"
        out = _sh(["systemctl", "start", "--no-block", unit])
        tail = ""
        if UPDATE_LOG.exists():
            tail = "\n".join(UPDATE_LOG.read_text(encoding="utf-8",
                                                  errors="replace").splitlines()[-15:])
        return jsonify({"started": unit, "systemctl": out, "log_tail": tail})

    @app.route("/api/admin/run", methods=["POST"])
    @requires_auth
    def admin_run():
        guard_write()
        data = request.get_json(silent=True) or {}
        try:
            argv = _build(str(data.get("cmd", "")), list(data.get("args") or []))
        except ValueError as exc:
            return jsonify({"error": str(exc)}), 400
        JOBS.mkdir(parents=True, exist_ok=True)
        # Журналы задач старше недели — в корзину; больше MAX_JOBS задач
        # одновременно не запускаем: каждая может тратить токены модели.
        now = time.time()
        for old in JOBS.glob("*.log"):
            if now - old.stat().st_mtime > 7 * 86400:
                old.unlink(missing_ok=True)
        running = _sh(["systemctl", "list-units", "--type=service", "--state=active,activating",
                       "--no-legend", "--plain", "jobsignal-job-*"])
        if len([ln for ln in running.splitlines() if ln.strip()]) >= MAX_JOBS:
            return jsonify({"error": f"уже идут {MAX_JOBS} задачи — дождись их"}), 429
        job = time.strftime("%m%d-%H%M%S-") + uuid.uuid4().hex[:6]
        log_path = JOBS / f"{job}.log"
        shell = (f"echo '$ {' '.join(shlex.quote(a) for a in argv[1:] if a)}'; "
                 f"{' '.join(shlex.quote(a) for a in argv)}; echo \"[код выхода $?]\"")
        # Отдельный юнит: задача переживает перезапуск кабинета и не держит
        # его процесс. Каталог и окружение — как у остальных служб.
        out = subprocess.run(
            ["systemd-run", f"--unit=jobsignal-job-{job}", "--collect", "--quiet",
             f"--property=WorkingDirectory={APP_DIR}",
             "--property=Nice=10", "--property=RuntimeMaxSec=3600",
             "/bin/sh", "-c", f"( {shell} ) > {shlex.quote(str(log_path))} 2>&1"],
            capture_output=True, text=True, timeout=20)
        if out.returncode:
            return jsonify({"error": (out.stdout + out.stderr).strip()}), 500
        return jsonify({"job": job})

    @app.route("/api/admin/jobs/<job>")
    @requires_auth
    def admin_job(job: str):
        if not re.fullmatch(r"\d{4}-\d{6}-[0-9a-f]{6}", job):
            abort(404)
        path = JOBS / f"{job}.log"
        if not path.exists():
            abort(404)
        running = _sh(["systemctl", "is-active", f"jobsignal-job-{job}.service"]) in (
            "active", "activating")
        text = path.read_text(encoding="utf-8", errors="replace")
        return jsonify({"job": job, "running": running, "output": text[-60000:]})

    # Диагностика: снимки и списки data-qa, которые оставляет автоотклик.
    FILE_DIRS = {"shots": APP_DIR / "data" / "hh_screenshots",
                 "choice": APP_DIR / "data" / "try_resume_choice",
                 "probe": APP_DIR / "data" / "probe_resume_choice"}

    @app.route("/api/admin/files/<kind>")
    @requires_auth
    def admin_files(kind: str):
        base = FILE_DIRS.get(kind)
        if base is None:
            abort(404)
        name = request.args.get("name", "")
        if not name:
            items = sorted(base.glob("*"), key=lambda p: p.stat().st_mtime,
                           reverse=True)[:int(request.args.get("n", "40"))] if base.exists() else []
            return jsonify([{"name": p.name, "size": p.stat().st_size} for p in items])
        if not re.fullmatch(r"[\w.-]{1,120}\.(png|jpg|txt|html|json)", name):
            abort(400)
        path = base / name
        if not path.is_file():
            abort(404)
        from flask import send_file
        return send_file(str(path), max_age=0)

    @app.route("/api/admin/logs")
    @requires_auth
    def admin_logs():
        unit = request.args.get("unit", "")
        if unit not in UNITS:
            return jsonify({"error": f"служба одна из: {', '.join(UNITS)}"}), 400
        n = request.args.get("n", "200")
        n = n if _INT.match(n) or n in ("1000", "2000") else "200"
        since = request.args.get("since", "")
        argv = ["journalctl", "-u", f"{unit}.service", "-n", n, "--no-pager",
                "-o", "short-iso"]
        if re.fullmatch(r"\d{1,3} (min|hour|day)s? ago", since):
            argv += ["--since", since]
        until = request.args.get("until", "")
        if re.fullmatch(r"\d{4}-\d\d-\d\d \d\d:\d\d(:\d\d)?", until):
            argv += ["--until", until]
        if re.fullmatch(r"\d{4}-\d\d-\d\d \d\d:\d\d(:\d\d)?", since):
            argv += ["--since", since]
        return jsonify({"unit": unit, "log": _sh(argv, timeout=30)})
