#!/bin/bash
# Сервис обработки нажатий кнопки «Получить сопроводительное»
# Запуск: bash deploy/setup_notify_polling.sh

set -e
PROJECT=/opt/jobsignal_local
PYTHON=$PROJECT/.venv/bin/python

echo "==> Создаём polling-сервис для кнопок бота"

cat > /etc/systemd/system/jobsignal-bot-polling.service << EOF
[Unit]
Description=JobSignal bot callback polling
After=network.target

[Service]
Type=simple
WorkingDirectory=$PROJECT
ExecStart=$PYTHON -c "
from dotenv import load_dotenv; load_dotenv('config/.env')
import logging; logging.basicConfig(level=logging.INFO, format='%(asctime)s %(message)s')
from jobsignal.agents.notify_bot import NotifyBot
NotifyBot().run_webhook_polling()
"
Restart=always
RestartSec=10
StandardOutput=journal
StandardError=journal

[Install]
WantedBy=multi-user.target
EOF

systemctl daemon-reload
systemctl enable --now jobsignal-bot-polling.service
echo "Polling-сервис запущен"
systemctl status jobsignal-bot-polling --no-pager | head -10
