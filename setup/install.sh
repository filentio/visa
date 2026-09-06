#!/bin/bash
set -e

echo "=============================="
echo " Установка окружения на VPS"
echo "=============================="

# --- 1. Обновление системы ---
echo "[1/7] Обновление системы..."
apt-get update -q && apt-get upgrade -y -q

# --- 2. Базовые пакеты ---
echo "[2/7] Установка базовых пакетов..."
apt-get install -y -q \
    git curl wget tmux htop \
    python3 python3-pip python3-venv \
    postgresql postgresql-contrib \
    redis-server \
    build-essential \
    fail2ban \
    ufw

# --- 3. Node.js 20 ---
echo "[3/7] Установка Node.js 20..."
curl -fsSL https://deb.nodesource.com/setup_20.x | bash -
apt-get install -y nodejs
echo "Node.js: $(node --version)"
echo "npm: $(npm --version)"

# --- 4. Claude Code ---
echo "[4/7] Установка Claude Code..."
npm install -g @anthropic-ai/claude-code
echo "Claude Code: $(claude --version 2>/dev/null || echo 'установлен')"

# --- 5. Python окружение для парсера ---
echo "[5/7] Настройка Python окружения..."
mkdir -p /opt/jobparser
python3 -m venv /opt/jobparser/venv
/opt/jobparser/venv/bin/pip install --upgrade pip -q
/opt/jobparser/venv/bin/pip install -q \
    playwright \
    aiohttp \
    asyncpg \
    python-dotenv \
    anthropic \
    celery[redis] \
    python-telegram-bot \
    beautifulsoup4 \
    lxml \
    pydantic

# --- 6. Playwright браузер ---
echo "[6/7] Установка Playwright + Chromium..."
/opt/jobparser/venv/bin/playwright install chromium
/opt/jobparser/venv/bin/playwright install-deps chromium

# --- 7. Безопасность ---
echo "[7/7] Настройка безопасности..."

# fail2ban
systemctl enable fail2ban
systemctl start fail2ban

# UFW - разрешаем SSH и закрываем лишнее
ufw --force reset
ufw default deny incoming
ufw default allow outgoing
ufw allow 22/tcp
ufw --force enable

# Отключаем вход по паролю (только ключи)
sed -i 's/#PasswordAuthentication yes/PasswordAuthentication no/' /etc/ssh/sshd_config
sed -i 's/PasswordAuthentication yes/PasswordAuthentication no/' /etc/ssh/sshd_config
systemctl reload ssh

echo ""
echo "=============================="
echo " Установка завершена!"
echo "=============================="
echo ""
echo "Следующие шаги:"
echo "1. Добавь API ключ Anthropic:"
echo "   echo 'export ANTHROPIC_API_KEY=sk-ant-XXXXX' >> ~/.bashrc && source ~/.bashrc"
echo ""
echo "2. Запусти Claude Code в tmux:"
echo "   tmux new -s claude && claude"
echo ""
echo "3. Настрой .env для парсера:"
echo "   cp /opt/jobparser/.env.example /opt/jobparser/.env && nano /opt/jobparser/.env"
echo ""
