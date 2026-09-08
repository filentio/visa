#!/bin/bash
# Добавляет сбор с hh.ru в systemd (каждые 4 часа)
# Запуск: bash deploy/setup_hh_collector.sh

set -e
PROJECT=/opt/jobsignal_local
PYTHON=$PROJECT/.venv/bin/python

echo "==> Создаём сервис и таймер hh.ru сборщика"

cat > /etc/systemd/system/jobsignal-hh.service << EOF
[Unit]
Description=JobSignal hh.ru collector
After=network.target

[Service]
Type=oneshot
WorkingDirectory=$PROJECT
ExecStart=$PYTHON run.py hh-collect
StandardOutput=journal
StandardError=journal
EOF

cat > /etc/systemd/system/jobsignal-hh.timer << EOF
[Unit]
Description=JobSignal hh.ru collection every 4 hours
Requires=jobsignal-hh.service

[Timer]
OnCalendar=*-*-* 06,10,14,18,22:00:00
Persistent=true

[Install]
WantedBy=timers.target
EOF

systemctl daemon-reload
systemctl enable --now jobsignal-hh.timer
echo "Таймер hh.ru сборщика включён"
systemctl list-timers | grep jobsignal
