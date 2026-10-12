#!/usr/bin/env bash
# Развернуть jobsignal на чистом сервере (Ubuntu 24.04) — или переехать.
#
# Зачем отдельный скрипт, а не setup_server.sh. Тот писался под первую версию
# и для работающей системы вреден: открывает порт дашборда наружу (его держим
# закрытым, ходим через SSH-туннель), пишет поверх правильных таймеров
# устаревшие (конвейер раз в 6 часов вместо раза в час) и ждёт ключ GigaChat.
#
# Два шага, и это важно:
#
#   bash bootstrap.sh            — ставит всё, но НИЧЕГО не запускает;
#   bash bootstrap.sh --enable   — проверяет настройки и включает таймеры.
#
# Почему не сразу. Включённые таймеры на сервере без резюме, без сессии hh и
# с пустым .env дадут час громких падений. А если старый сервер ещё жив,
# два сервера начнут откликаться параллельно, и бот уведомлений сломается:
# Telegram отдаёт обновления только одному получателю на токен.
#
# Повторный запуск безопасен.
#
# Запуск на новом сервере (одной строкой):
#   curl -fsSL https://raw.githubusercontent.com/filentio/visa/claude/gracious-goldberg-1dyvtg/jobsignal/deploy/bootstrap.sh | bash
set -euo pipefail

REPO_URL="https://github.com/filentio/visa.git"
BRANCH="${JOBSIGNAL_BRANCH:-claude/gracious-goldberg-1dyvtg}"
SRC=/root/visa
APP=/opt/jobsignal_local
PY="$APP/.venv/bin/python"

[ "$(id -u)" -eq 0 ] || { echo "запусти от root"; exit 1; }

enable_mode=0
[ "${1:-}" = "--enable" ] && enable_mode=1

if [ "$enable_mode" -eq 0 ]; then
  echo "==> 1/7 swap 2 ГБ"
  # Chromium в пике берёт ~660 МБ. Без свопа один прогон браузера рядом с
  # конвейером может уронить процесс по памяти.
  if ! swapon --show | grep -q .; then
    fallocate -l 2G /swapfile 2>/dev/null || dd if=/dev/zero of=/swapfile bs=1M count=2048
    chmod 600 /swapfile && mkswap /swapfile >/dev/null && swapon /swapfile
    grep -q '^/swapfile' /etc/fstab || echo '/swapfile none swap sw 0 0' >> /etc/fstab
  fi

  echo "==> 2/7 системные пакеты"
  export DEBIAN_FRONTEND=noninteractive
  apt-get update -qq
  apt-get install -y -qq python3-venv python3-pip git curl ufw >/dev/null

  echo "==> 3/7 код"
  if [ -d "$SRC/.git" ]; then
    git -C "$SRC" fetch -q origin "$BRANCH"
    git -C "$SRC" checkout -q "$BRANCH"
    git -C "$SRC" pull -q origin "$BRANCH"
  else
    git clone -q -b "$BRANCH" "$REPO_URL" "$SRC"
  fi
  mkdir -p "$APP" "$APP/config" "$APP/data" /root/autoapply
  # Код копируем, а config/ и data/ не трогаем: там секреты, резюме и база,
  # которые в репозиторий не попадают и при обновлении не должны затираться.
  cp -r "$SRC/jobsignal/jobsignal" "$SRC/jobsignal/tools" "$SRC/jobsignal/deploy" "$APP/"
  cp "$SRC/jobsignal/run.py" "$SRC/jobsignal/run_polling.py" "$SRC/jobsignal/requirements.txt" "$APP/"
  for f in channels.yaml profiles.yaml .env.example; do
    cp "$SRC/jobsignal/config/$f" "$APP/config/$f"
  done

  echo "==> 4/7 окружение Python (несколько минут)"
  [ -d "$APP/.venv" ] || python3 -m venv "$APP/.venv"
  "$APP/.venv/bin/pip" install -q --upgrade pip
  "$APP/.venv/bin/pip" install -q -r "$APP/requirements.txt"

  echo "==> 5/7 браузер для кабинета hh"
  "$APP/.venv/bin/playwright" install --with-deps chromium >/dev/null

  echo "==> 6/7 база и каналы"
  [ -f "$APP/config/.env" ] || { cp "$APP/config/.env.example" "$APP/config/.env"; chmod 600 "$APP/config/.env"; }
  cd "$APP"
  "$PY" run.py initdb
  "$PY" run.py seed-channels

  echo "==> 7/7 службы и фаервол"
  cp "$APP"/deploy/systemd/*.service "$APP"/deploy/systemd/*.timer /etc/systemd/system/
  systemctl daemon-reload
  # Только SSH. Дашборд слушает 5000, но снаружи закрыт: смотреть через
  # туннель  ssh -L 5000:localhost:5000 root@<сервер>
  ufw allow OpenSSH >/dev/null
  ufw deny 5000/tcp >/dev/null
  yes | ufw enable >/dev/null

  cat <<MSG

================ УСТАНОВЛЕНО, НО НЕ ЗАПУЩЕНО ================
Осталось руками (подробно — в ответе в чате):
  1. nano $APP/config/.env          — ключи и настройки
  2. резюме в $APP/config/          — cv_ai_pm.md, cv_cpo.md, cv_pm.md,
                                      master_cv.pdf, cv_stories.md
  3. сессия hh  → /root/autoapply/state.json  (tools/hh_login_local.py
                                      на своём компьютере, через тот же прокси)
  4. сессия тг  → $PY tools/tg_login.py
Потом:  bash $SRC/jobsignal/deploy/bootstrap.sh --enable
=============================================================
MSG
  exit 0
fi

# ---------------------------- --enable ----------------------------------
cd "$APP"
fail=0
need() { grep -qE "^$1=.+" config/.env || { echo "  !! в config/.env не задан $1"; fail=1; }; }
echo "==> проверка настроек"
need ANTHROPIC_API_KEY
need NOTIFY_BOT_TOKEN
need NOTIFY_CHAT_ID
need DASHBOARD_PASS
for f in cv_ai_pm.md cv_cpo.md cv_pm.md; do
  [ -s "config/$f" ] || { echo "  !! нет config/$f — матчер не сможет оценивать"; fail=1; }
done
[ "$fail" -eq 0 ] || { echo "исправь и запусти --enable снова"; exit 1; }

# Это не блокирует запуск: без них работают сбор, разбор и оценка. Но
# отправка на hh и сторож телеграма упадут громко, так что лучше знать.
[ -s config/master_cv.pdf ]      || echo "  предупреждение: нет master_cv.pdf — письма к hh без текста резюме"
[ -s config/cv_stories.md ]      || echo "  предупреждение: нет cv_stories.md — сторож не будет готовить черновики"
[ -s /root/autoapply/state.json ] || echo "  предупреждение: нет сессии hh — отправка и учёт ответов не заработают"
[ -s data/tg_inbox.session ]     || echo "  предупреждение: нет сессии тг — сторож и проверка дублей не заработают"
grep -qE '^HH_PROXY=.+' config/.env || echo "  предупреждение: HH_PROXY пуст — кабинет hh с адреса сервера закрыт ddos-guard"

echo "==> включаю службы"
systemctl enable --now jobsignal-bot.service jobsignal-dashboard.service
systemctl enable --now jobsignal-pipeline.timer jobsignal-hh-apply.timer \
  jobsignal-replies.timer jobsignal-tgwatch.timer jobsignal-backup.timer
systemctl list-timers 'jobsignal*' --no-pager
echo
echo "Первый прогон конвейера сразу:  systemctl start jobsignal-pipeline.service"
