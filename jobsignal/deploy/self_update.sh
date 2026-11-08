#!/bin/bash
# Обновление сервера с ветки репозитория. Запускают jobsignal-update.service
# (таймер раз в 10 минут и кнопка /api/admin/update в кабинете).
#
#   * код берётся из /root/visa (git pull --ff-only по текущей ветке);
#   * сначала проверка: новый код импортируется — иначе ничего не меняем;
#   * копируется только код (jobsignal/, tools/, deploy/, run.py ...):
#     config/ и data/ на сервере — секреты, профили, база — не трогаются;
#   * пока копируем, держим замок hh: конвейер и отклик не стартуют посреди
#     подмены файлов;
#   * перезапуск бота и кабинета, сообщение в телеграм с тем, что обновилось.
#
#   --force  — разложить и перезапустить, даже если новых коммитов нет.
set -uo pipefail

main() {

SRC=${JOBSIGNAL_SRC:-/root/visa}
APP=${JOBSIGNAL_APP:-/opt/jobsignal_local}
PY=$APP/.venv/bin/python
STATE=$APP/data/update_state.json
LOG=$APP/data/update.log
FORCE=${1:-}

mkdir -p "$APP/data"
exec > >(tee -a "$LOG") 2>&1
echo "=== $(date -u '+%d.%m.%Y %H:%M:%S UTC') обновление"

state() {  # ok|fail|same  сообщение
  "$PY" - "$1" "$2" "$STATE" <<'PY'
import json, sys, time
ok, msg, path = sys.argv[1], sys.argv[2], sys.argv[3]
json.dump({"result": ok, "message": msg,
           "at": time.strftime("%d.%m.%Y %H:%M UTC", time.gmtime())},
          open(path, "w"), ensure_ascii=False)
PY
}

notify() {
  ( cd "$APP" && "$PY" - "$1" <<'PY'
import sys
from dotenv import load_dotenv
load_dotenv("config/.env")
from jobsignal.agents.notify_bot import _esc, _send
_send(_esc(sys.argv[1]))
PY
  ) || echo "сообщение в телеграм не ушло"
}

cd "$SRC" || { echo "нет $SRC"; exit 1; }
BRANCH=$(git rev-parse --abbrev-ref HEAD)
OLD=$(git rev-parse HEAD)
if ! git fetch -q origin "$BRANCH"; then
  echo "git fetch не прошёл"; state fail "git fetch не прошёл"; exit 1
fi
NEW=$(git rev-parse "origin/$BRANCH")
if [ "$OLD" = "$NEW" ] && [ "$FORCE" != "--force" ]; then
  echo "новых коммитов нет ($(git log -1 --format=%h))"
  exit 0
fi
if ! git merge -q --ff-only "origin/$BRANCH"; then
  echo "в $SRC локальные изменения — обновление без перемотки невозможно"
  state fail "в /root/visa локальные изменения, git merge --ff-only не прошёл"
  notify "⚠️ Обновление не прошло: в /root/visa на сервере локальные изменения."
  exit 1
fi
CHANGES=$(git log --format='• %s' "$OLD..$NEW" | head -15)

# 1. Проверка: новый код импортируется (из исходников, до раскладки).
STAGE=$SRC/jobsignal
if ! ( cd "$APP" && PYTHONPATH="$STAGE" "$PY" -c "
import sys; sys.path.insert(0, '$STAGE')
import jobsignal.dashboard.app, jobsignal.agents.hh_apply, jobsignal.agents.notify_bot
import jobsignal.letters, jobsignal.llm
" ); then
  echo "новый код не импортируется — раскладку отменяю"
  git reset -q --keep "$OLD"
  state fail "новый код не прошёл проверку импорта, оставлен $(git log -1 --format=%h)"
  notify "⚠️ Обновление отменено: новый код не запускается. Остался прежний."
  exit 1
fi

# 2. Раскладка под замком hh (ждём до 25 минут, пока закончится конвейер).
exec 9>"${JOBSIGNAL_HH_LOCK:-/run/jobsignal-hh.lock}"
if ! flock -w 1500 9; then
  echo "замок hh занят 25 минут — попробую в следующий раз"
  git reset -q --keep "$OLD"
  exit 0
fi
REQ_CHANGED=$(git diff --name-only "$OLD" "$NEW" -- jobsignal/requirements.txt)
UNITS_CHANGED=$(git diff --name-only "$OLD" "$NEW" -- jobsignal/deploy/systemd)
cp -r "$STAGE/jobsignal" "$STAGE/tools" "$STAGE/deploy" "$APP/"
cp "$STAGE/run.py" "$STAGE/run_polling.py" "$STAGE/requirements.txt" "$APP/"
for f in channels.yaml .env.example; do cp "$STAGE/config/$f" "$APP/config/$f"; done
if [ -n "$REQ_CHANGED" ]; then
  echo "requirements.txt изменился — ставлю зависимости"
  "$APP/.venv/bin/pip" install -q -r "$APP/requirements.txt"
fi
if [ -n "$UNITS_CHANGED" ] || [ "$FORCE" = "--force" ]; then
  cp "$APP"/deploy/systemd/*.service "$APP"/deploy/systemd/*.timer /etc/systemd/system/
  systemctl daemon-reload
fi
flock -u 9

# 3. Перезапуск долгоживущих служб (таймерные подхватят код сами).
systemctl restart jobsignal-bot.service
systemctl restart jobsignal-dashboard.service
sleep 4
BAD=""
for u in jobsignal-bot jobsignal-dashboard; do
  systemctl is-active --quiet "$u.service" || BAD="$BAD $u"
done
SHORT=$(git log -1 --format=%h)
if [ -n "$BAD" ]; then
  state fail "код $SHORT разложен, но не поднялись:$BAD"
  notify "⚠️ Обновлено до $SHORT, но не поднялись службы:$BAD"
  exit 1
fi
state ok "код $SHORT"
notify "🔄 Система обновлена до $SHORT
$CHANGES"
echo "готово: $SHORT"
}

main "$@"
exit $?
