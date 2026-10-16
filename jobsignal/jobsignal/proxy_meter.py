"""Учёт трафика через платный прокси и предупреждение, пока пакет не кончился.

Зачем. 15.10 пакет прокси кончился молча. Провайдер в таком случае не
отвечает отказом, а держит соединение, поэтому система видела только
таймауты: «страница не открылась (сеть или прокси)» каждые 5 минут, и
сколько часов уже не уходят отклики и не идёт поиск hh, было не понять.
До этого трафик считал только поиск hh, и то в пределах одного прогона;
браузер (автоотклик, учёт ответов, подбор под резюме, проба сессии) не
считался совсем.

Как считаем. Браузер: на каждом контексте с прокси слушаем requestfinished
и складываем размеры заголовков и тела ответа по проводу (request.sizes()).
Поиск: размер ответа requests по Content-Length. Это занижение: TLS,
CONNECT и повторные передачи провайдер тоже берёт в счёт, поэтому цифру
можно поправить множителем PROXY_METER_FACTOR, сверив с кабинетом
провайдера за пару дней.

Где лежит. data/proxy_traffic.json: байты по дням и источникам, сумма с
последнего пополнения и какие предупреждения уже ушли. Пишут несколько
процессов (прогон, автоотклик, учёт ответов), поэтому запись под fcntl.

Настройки в .env:
  PROXY_PACKAGE_GB   объём пакета; пусто — только учёт, без предупреждений
  PROXY_WARN_PCT     предупредить, когда осталось меньше, % (20)
  PROXY_METER_FACTOR поправка к замеру (1.0)

После пополнения: run.py proxy-traffic --topup [ГБ].
"""
from __future__ import annotations

import atexit
import fcntl
import json
import logging
import os
from contextlib import contextmanager
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

log = logging.getLogger("jobsignal")

KEEP_DAYS = 31
# Копим в памяти и сбрасываем в файл при выходе процесса или когда набралось
# столько: на каждый запрос браузера открывать файл незачем.
FLUSH_BYTES = 2 * 2**20

_pending: dict[str, int] = {}
_pending_total = 0
_atexit_set = False


def _path() -> Path:
    return Path(os.environ.get("PROXY_TRAFFIC_FILE", "data/proxy_traffic.json"))


def _factor() -> float:
    try:
        return float(os.environ.get("PROXY_METER_FACTOR", "1.0"))
    except ValueError:
        return 1.0


def package_bytes() -> int | None:
    raw = (os.environ.get("PROXY_PACKAGE_GB") or "").strip().replace(",", ".")
    if not raw:
        return None
    try:
        gb = float(raw)
    except ValueError:
        log.warning("[proxy] PROXY_PACKAGE_GB=%r не число — предупреждений не будет", raw)
        return None
    return int(gb * 2**30) if gb > 0 else None


def warn_pct() -> float:
    try:
        return float(os.environ.get("PROXY_WARN_PCT", "20"))
    except ValueError:
        return 20.0


@contextmanager
def _locked():
    """Прочитать и переписать файл учёта под замком."""
    path = _path()
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path.with_suffix(".lock"), "w") as lock:
        fcntl.flock(lock, fcntl.LOCK_EX)
        try:
            data = json.loads(path.read_text())
        except (OSError, ValueError):
            data = {}
        data.setdefault("days", {})
        data.setdefault("since_topup", 0)
        data.setdefault("warned", [])
        yield data
        cutoff = (date.today() - timedelta(days=KEEP_DAYS)).isoformat()
        data["days"] = {d: v for d, v in data["days"].items() if d >= cutoff}
        tmp = path.with_suffix(".tmp")
        tmp.write_text(json.dumps(data, ensure_ascii=False, indent=1))
        tmp.replace(path)


def add(source: str, nbytes: int) -> None:
    """Учесть трафик. Дёшево: только память, файл — при сбросе."""
    global _pending_total, _atexit_set
    if nbytes <= 0:
        return
    _pending[source] = _pending.get(source, 0) + int(nbytes)
    _pending_total += int(nbytes)
    if not _atexit_set:
        atexit.register(flush)
        _atexit_set = True
    if _pending_total >= FLUSH_BYTES:
        flush()


def flush() -> None:
    """Сбросить накопленное в файл и проверить остаток пакета."""
    global _pending_total
    if not _pending:
        return
    batch = dict(_pending)
    _pending.clear()
    _pending_total = 0
    factor = _factor()
    try:
        with _locked() as data:
            day = data["days"].setdefault(date.today().isoformat(), {})
            for src, n in batch.items():
                n = int(n * factor)
                day[src] = day.get(src, 0) + n
                data["since_topup"] += n
            warning = _due_warning(data)
    except OSError as exc:
        log.warning("[proxy] учёт трафика не записан: %s", exc)
        return
    if warning:
        _notify(warning)


def _due_warning(data: dict) -> str | None:
    """Текст предупреждения, если пора и оно ещё не уходило с пополнения."""
    pkg = package_bytes()
    if not pkg:
        return None
    left = pkg - data["since_topup"]
    left_pct = 100 * left / pkg
    # Две ступени: заранее и в последний момент. Каждая уходит один раз.
    for level in (warn_pct(), 5.0):
        if left_pct < level and level not in data["warned"]:
            data["warned"].append(level)
            days = _days_left(data, left)
            return (f"⚠️ <b>Прокси: осталось ~{max(left, 0) / 2**20:.0f} МБ "
                    f"из {pkg / 2**30:g} ГБ ({max(left_pct, 0):.0f}%)</b>\n"
                    + (f"При нынешнем расходе хватит примерно на {days}.\n"
                       if days else "")
                    + "Когда пакет кончится, перестанут уходить отклики и "
                      "идти поиск hh. После пополнения на сервере: "
                      "<code>run.py proxy-traffic --topup</code>")
    return None


def _avg_per_day(data: dict, days: int = 3) -> float:
    """Средний расход за последние полные дни (сегодняшний неполный не берём)."""
    today = date.today()
    vals = [sum(data["days"].get((today - timedelta(days=i)).isoformat(), {}).values())
            for i in range(1, days + 1)]
    vals = [v for v in vals if v]
    return sum(vals) / len(vals) if vals else 0.0


def _days_left(data: dict, left: int) -> str:
    avg = _avg_per_day(data)
    if avg <= 0 or left <= 0:
        return ""
    d = left / avg
    return f"{d:.0f} дн." if d >= 1 else f"{d * 24:.0f} ч"


def _notify(text: str) -> None:
    try:
        from jobsignal.agents.notify_bot import _send
        _send(text)
    except Exception as exc:  # noqa: BLE001 — предупреждение не должно ронять прогон
        log.warning("[proxy] предупреждение о пакете не отправлено: %s", exc)
    log.warning("[proxy] %s", text.replace("<b>", "").replace("</b>", ""))


def topup(package_gb: float | None = None) -> dict:
    """Пакет пополнен: обнулить счётчик и разрешить предупреждения снова."""
    flush()
    with _locked() as data:
        data["since_topup"] = 0
        data["warned"] = []
        data["topup_at"] = datetime.now(timezone.utc).isoformat()
        if package_gb:
            data["package_gb_note"] = package_gb
        return dict(data)


def summary() -> dict:
    """То же, что report(), но данными — для кабинета."""
    flush()
    with _locked() as data:
        snap = json.loads(json.dumps(data))
    today = snap["days"].get(date.today().isoformat(), {})
    pkg = package_bytes()
    used = snap["since_topup"]
    left = (pkg - used) if pkg else None
    return {
        "today_mb": round(sum(today.values()) / 2**20, 1),
        "today_by_source": {k: round(v / 2**20, 1) for k, v in
                            sorted(today.items(), key=lambda kv: -kv[1])},
        "avg_day_mb": round(_avg_per_day(snap) / 2**20, 1),
        "since_topup_mb": round(used / 2**20, 1),
        "topup_at": snap.get("topup_at"),
        "package_gb": pkg / 2**30 if pkg else None,
        "left_mb": round(max(left, 0) / 2**20) if left is not None else None,
        "left_pct": round(100 * max(left, 0) / pkg) if pkg else None,
        "days_left": _days_left(snap, left) if left is not None else "",
        "days": {d: round(sum(v.values()) / 2**20, 1)
                 for d, v in sorted(snap["days"].items())[-7:]},
    }


def report() -> str:
    """Сводка для человека: по дням и источникам, остаток, прогноз."""
    flush()
    with _locked() as data:
        snapshot = json.loads(json.dumps(data))
    lines = []
    days = sorted(snapshot["days"])[-7:]
    if not days:
        lines.append("Учёт пуст: через прокси ещё ничего не ходило после установки.")
    for d in days:
        src = snapshot["days"][d]
        total = sum(src.values())
        parts = ", ".join(f"{k} {v / 2**20:.1f}"
                          for k, v in sorted(src.items(), key=lambda kv: -kv[1]))
        lines.append(f"{d}: {total / 2**20:7.1f} МБ  ({parts})")
    avg = _avg_per_day(snapshot)
    if avg:
        lines.append(f"Средний расход за 3 полных дня: {avg / 2**20:.0f} МБ/сут")
    used = snapshot["since_topup"]
    since = snapshot.get("topup_at", "")[:16].replace("T", " ")
    lines.append(f"С пополнения{(' ' + since + ' UTC') if since else ''}: "
                 f"{used / 2**20:.0f} МБ")
    pkg = package_bytes()
    if pkg:
        left = pkg - used
        days_left = _days_left(snapshot, left)
        lines.append(f"Пакет {pkg / 2**30:g} ГБ, осталось ~{max(left, 0) / 2**20:.0f} МБ"
                     + (f", хватит на {days_left}" if days_left else ""))
    else:
        lines.append("PROXY_PACKAGE_GB не задан — остаток не считается, "
                     "предупреждений не будет")
    factor = _factor()
    if factor != 1.0:
        lines.append(f"Поправка к замеру: ×{factor:g}")
    return "\n".join(lines)


def meter_context(ctx, source: str) -> None:
    """Считать трафик браузерного контекста, если он ходит через прокси."""
    from jobsignal.hh_session import proxy_url
    if not proxy_url():
        return

    async def _done(request) -> None:
        try:
            sizes = await request.sizes()
        except Exception:  # noqa: BLE001 — контекст уже закрыт
            return
        add(source, sizes.get("responseBodySize", 0)
            + sizes.get("responseHeadersSize", 0)
            + sizes.get("requestHeadersSize", 0)
            + sizes.get("requestBodySize", 0))

    ctx.on("requestfinished", _done)
