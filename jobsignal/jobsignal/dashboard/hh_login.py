"""Страница «Вход в hh» (/hh-login): управление браузером на сервере.

Сам браузер — tools/hh_login_remote.py, отдельный юнит jobsignal-hh-login
(systemd-run): переживает перезапуск кабинета и не держит его процесс.
Обмен — файлами в data/hh_login/ (см. описание в том скрипте).
"""
from __future__ import annotations

import json
import subprocess
from pathlib import Path

from flask import abort, jsonify, render_template, request, send_file

APP_DIR = Path(__file__).resolve().parent.parent.parent
DIR = APP_DIR / "data" / "hh_login"
UNIT = "jobsignal-hh-login"
KEYS = ("Enter", "Backspace", "Tab", "Escape", "ArrowDown", "ArrowUp", "Delete")


def _running() -> bool:
    out = subprocess.run(["systemctl", "is-active", f"{UNIT}.service"],
                         capture_output=True, text=True)
    return out.stdout.strip() in ("active", "activating")


def _push(cmd: dict) -> None:
    DIR.mkdir(parents=True, exist_ok=True)
    with open(DIR / "commands.jsonl", "a", encoding="utf-8") as fh:
        fh.write(json.dumps(cmd, ensure_ascii=False) + "\n")


def register(app, requires_auth, same_origin) -> None:

    def guard_write():
        if not same_origin():
            abort(403)

    @app.route("/hh-login")
    @requires_auth
    def hh_login_page():
        return render_template("hh_login.html")

    @app.route("/api/hh-login/start", methods=["POST"])
    @requires_auth
    def hh_login_start():
        guard_write()
        if _running():
            return jsonify({"ok": True, "message": "браузер уже запущен"})
        DIR.mkdir(parents=True, exist_ok=True)
        for f in ("screen.jpg", "status.json", "commands.jsonl"):
            (DIR / f).unlink(missing_ok=True)
        py = str(APP_DIR / ".venv" / "bin" / "python")
        out = subprocess.run(
            ["systemd-run", f"--unit={UNIT}", "--collect", "--quiet",
             f"--property=WorkingDirectory={APP_DIR}", "--property=RuntimeMaxSec=1200",
             py, "tools/hh_login_remote.py"],
            capture_output=True, text=True, timeout=20)
        if out.returncode:
            return jsonify({"ok": False, "message": (out.stdout + out.stderr).strip()}), 500
        return jsonify({"ok": True, "message": "запускаю браузер"})

    @app.route("/api/hh-login/status")
    @requires_auth
    def hh_login_status():
        try:
            st = json.loads((DIR / "status.json").read_text(encoding="utf-8"))
        except (OSError, ValueError):
            st = {"state": "idle", "message": ""}
        st["running"] = _running()
        if st.get("state") in ("starting", "login", "saving") and not st["running"]:
            st["state"], st["message"] = "failed", "браузер завершился — запустите вход заново"
        return jsonify(st)

    @app.route("/api/hh-login/screen")
    @requires_auth
    def hh_login_screen():
        path = DIR / "screen.jpg"
        if not path.exists():
            abort(404)
        resp = send_file(str(path), mimetype="image/jpeg", max_age=0)
        resp.headers["Cache-Control"] = "no-store"
        return resp

    @app.route("/api/hh-login/input", methods=["POST"])
    @requires_auth
    def hh_login_input():
        guard_write()
        d = request.get_json(silent=True) or {}
        t = d.get("t")
        if t == "click":
            try:
                x, y = float(d["x"]), float(d["y"])
            except (KeyError, TypeError, ValueError):
                abort(400)
            if not (0 <= x <= 1280 and 0 <= y <= 800):
                abort(400)
            _push({"t": "click", "x": x, "y": y})
        elif t == "type":
            text = str(d.get("text", ""))[:200]
            if not text:
                abort(400)
            _push({"t": "type", "text": text})
        elif t == "key" and d.get("key") in KEYS:
            _push({"t": "key", "key": d["key"]})
        elif t == "scroll":
            try:
                _push({"t": "scroll", "dy": max(-2000, min(2000, int(d.get("dy", 0))))})
            except (TypeError, ValueError):
                abort(400)
        elif t in ("reload", "login", "stop"):
            _push({"t": t})
        else:
            abort(400)
        return jsonify({"ok": True})
