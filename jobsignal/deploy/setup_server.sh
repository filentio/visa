#!/usr/bin/env bash
# Установка JobSignal Local на сервер (Ubuntu 22.04/24.04).
# Запускать из корня проекта от root:  sudo bash deploy/setup_server.sh
# Идемпотентно: можно гонять повторно.
set -euo pipefail

PROJECT="$(cd "$(dirname "$0")/.." && pwd)"
PYBIN="$PROJECT/.venv/bin/python"
PORT="${DASHBOARD_PORT:-5000}"

echo "==> Проект: $PROJECT"
[ "$(id -u)" -eq 0 ] || { echo "Запусти от root (sudo)"; exit 1; }

echo "==> 1/6 swap 2G (страховка для 1 ГБ RAM)"
if ! swapon --show | grep -q .; then
  fallocate -l 2G /swapfile || dd if=/dev/zero of=/swapfile bs=1M count=2048
  chmod 600 /swapfile && mkswap /swapfile && swapon /swapfile
  grep -q '/swapfile' /etc/fstab || echo '/swapfile none swap sw 0 0' >> /etc/fstab
  echo "   swap включён"
else
  echo "   swap уже есть — пропускаю"
fi

echo "==> 2/6 системные пакеты"
export DEBIAN_FRONTEND=noninteractive
apt-get update -qq
apt-get install -y -qq python3-venv python3-pip git curl >/dev/null

echo "==> 3/6 виртуальное окружение и зависимости"
[ -d "$PROJECT/.venv" ] || python3 -m venv "$PROJECT/.venv"
"$PROJECT/.venv/bin/pip" install -q --upgrade pip
"$PROJECT/.venv/bin/pip" install -q -r "$PROJECT/requirements.txt"

echo "==> 4/6 проверка .env"
if [ ! -f "$PROJECT/config/.env" ]; then
  cp "$PROJECT/config/.env.example" "$PROJECT/config/.env"
  echo "   !! создан config/.env из шаблона — впиши GIGACHAT_AUTH_KEY и DASHBOARD_PASS, потом перезапусти сервисы"
fi
"$PYBIN" "$PROJECT/run.py" initdb

echo "==> 5/6 systemd: дашборд (всегда) + конвейер (по таймеру)"
cat > /etc/systemd/system/jobsignal-dashboard.service <<UNIT
[Unit]
Description=JobSignal dashboard
After=network.target
[Service]
WorkingDirectory=$PROJECT
ExecStart=$PYBIN $PROJECT/run.py serve
Restart=always
RestartSec=5
[Install]
WantedBy=multi-user.target
UNIT

cat > /etc/systemd/system/jobsignal-pipeline.service <<UNIT
[Unit]
Description=JobSignal pipeline (collect, parse, dedup, match)
[Service]
Type=oneshot
WorkingDirectory=$PROJECT
ExecStart=$PYBIN $PROJECT/run.py pipeline
UNIT

cat > /etc/systemd/system/jobsignal-pipeline.timer <<UNIT
[Unit]
Description=Run JobSignal pipeline every 6h
[Timer]
OnBootSec=5min
OnUnitActiveSec=6h
Persistent=true
[Install]
WantedBy=timers.target
UNIT

systemctl daemon-reload
systemctl enable --now jobsignal-dashboard.service
systemctl enable --now jobsignal-pipeline.timer

echo "==> 6/6 фаервол"
if command -v ufw >/dev/null; then
  ufw allow OpenSSH >/dev/null 2>&1 || ufw allow 22/tcp >/dev/null 2>&1 || true
  ufw allow "${PORT}/tcp" >/dev/null 2>&1 || true
  yes | ufw enable >/dev/null 2>&1 || true
  echo "   ufw: открыты SSH и порт $PORT"
fi

IP="$(curl -s --max-time 5 ifconfig.me || echo '<IP-сервера>')"
echo
echo "================= ГОТОВО ================="
echo "Дашборд:   http://$IP:$PORT"
echo "Логи дашборда:   journalctl -u jobsignal-dashboard -f"
echo "Логи конвейера:  journalctl -u jobsignal-pipeline -f"
echo "Запустить сбор сейчас: systemctl start jobsignal-pipeline.service"
echo "После правки config/.env: systemctl restart jobsignal-dashboard"
echo "=========================================="
