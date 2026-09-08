#!/bin/bash
# Уведомления о новых вакансиях каждые 10 минут
# Запуск: bash deploy/setup_notify.sh

set -e
PROJECT=/opt/jobsignal_local
PYTHON=$PROJECT/.venv/bin/python

echo "==> Создаём сервис и таймер уведомлений"

cat > /etc/systemd/system/jobsignal-notify.service << EOF
[Unit]
Description=JobSignal new vacancy notifications
After=network.target

[Service]
Type=oneshot
WorkingDirectory=$PROJECT
ExecStart=$PYTHON run.py notify-check
StandardOutput=journal
StandardError=journal
EOF

cat > /etc/systemd/system/jobsignal-notify.timer << EOF
[Unit]
Description=JobSignal notify every 10 minutes
Requires=jobsignal-notify.service

[Timer]
OnCalendar=*:0/10
Persistent=true

[Install]
WantedBy=timers.target
EOF

systemctl daemon-reload
systemctl enable --now jobsignal-notify.timer
echo "Таймер уведомлений включён (каждые 10 минут)"
systemctl list-timers | grep jobsignal
