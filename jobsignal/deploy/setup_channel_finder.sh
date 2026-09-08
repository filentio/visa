#!/bin/bash
# Добавляет ежесуточный поиск каналов в systemd
# Запуск: bash deploy/setup_channel_finder.sh

set -e
PROJECT=/opt/jobsignal_local
PYTHON=$PROJECT/.venv/bin/python

echo "==> Создаём таймер поиска каналов (раз в сутки)"

cat > /etc/systemd/system/jobsignal-channels.service << EOF
[Unit]
Description=JobSignal channel finder
After=network.target

[Service]
Type=oneshot
WorkingDirectory=$PROJECT
ExecStart=$PYTHON run.py find-channels
StandardOutput=journal
StandardError=journal
EOF

cat > /etc/systemd/system/jobsignal-channels.timer << EOF
[Unit]
Description=JobSignal daily channel search
Requires=jobsignal-channels.service

[Timer]
OnCalendar=daily
Persistent=true
RandomizedDelaySec=3600

[Install]
WantedBy=timers.target
EOF

systemctl daemon-reload
systemctl enable --now jobsignal-channels.timer
echo "Таймер поиска каналов включён"
systemctl list-timers | grep jobsignal
