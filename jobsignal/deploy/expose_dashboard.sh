#!/usr/bin/env bash
# Открыть кабинет в интернет по HTTPS, без SSH-туннеля.
#
#   bash deploy/expose_dashboard.sh                 # адрес вида 5-187-4-252.sslip.io
#   bash deploy/expose_dashboard.sh my.domain.ru    # свой домен (A-запись на сервер)
#   bash deploy/expose_dashboard.sh --new-pass      # заодно сменить пароль кабинета
#
# Как устроено. Дашборд по-прежнему слушает только 127.0.0.1:5000, порт 5000
# снаружи закрыт. Перед ним встаёт Caddy: принимает 443, сам получает и
# продлевает сертификат Let's Encrypt и передаёт запросы дашборду. Своего
# домена не нужно: sslip.io отдаёт адрес сервера по имени вида 5-187-4-252,
# и сертификат на такое имя выдаётся как на обычный домен.
#
# Защита: HTTPS (пароль не идёт открытым текстом), пароль из DASHBOARD_PASS
# (не короче 16 знаков; если короче или пуст — скрипт заменит его на
# случайный и покажет один раз на этом экране), после 10 неверных паролей
# адрес закрывается на 15 минут, страницы не индексируются поисковиками.
set -euo pipefail

APP=/opt/jobsignal_local
ENV="$APP/config/.env"
[ "$(id -u)" -eq 0 ] || { echo "запусти от root"; exit 1; }
[ -f "$ENV" ] || { echo "нет $ENV — сначала bootstrap.sh"; exit 1; }

# ── адрес ────────────────────────────────────────────────────────────────────
HOST=""
FORCE_NEW=0
for a in "$@"; do
  case "$a" in
    --new-pass) FORCE_NEW=1 ;;
    *) HOST="$a" ;;
  esac
done
if [ -z "$HOST" ]; then
  IP=$(curl -4 -fsS --max-time 10 https://api.ipify.org || hostname -I | awk '{print $1}')
  [ -n "$IP" ] || { echo "не удалось узнать внешний адрес сервера"; exit 1; }
  HOST="${IP//./-}.sslip.io"
fi
echo "==> адрес кабинета: https://$HOST"

# ── пароль ───────────────────────────────────────────────────────────────────
# Как читает python-dotenv: хвост « # комментарий» и пробелы — не часть пароля.
CUR=$(grep -E '^DASHBOARD_PASS=' "$ENV" | tail -1 | cut -d= -f2- \
      | sed -E 's/[[:space:]]+#.*$//; s/^[[:space:]]+//; s/[[:space:]]+$//' | tr -d "\"'" || true)
NEWPASS=""
if [ "$FORCE_NEW" -eq 1 ] || [ "${#CUR}" -lt 16 ]; then
  NEWPASS=$(openssl rand -base64 24 | tr -d '/+=' | cut -c1-20)
  if grep -qE '^DASHBOARD_PASS=' "$ENV"; then
    sed -i "s|^DASHBOARD_PASS=.*|DASHBOARD_PASS=$NEWPASS|" "$ENV"
  else
    echo "DASHBOARD_PASS=$NEWPASS" >> "$ENV"
  fi
fi
grep -qE '^DASHBOARD_USER=' "$ENV" || echo "DASHBOARD_USER=admin" >> "$ENV"
USER_NAME=$(grep -E '^DASHBOARD_USER=' "$ENV" | tail -1 | cut -d= -f2- | tr -d "\"'")
chmod 600 "$ENV"

# ── Caddy ────────────────────────────────────────────────────────────────────
if ! command -v caddy >/dev/null; then
  echo "==> ставлю Caddy"
  export DEBIAN_FRONTEND=noninteractive
  apt-get update -qq
  if ! apt-get install -y -qq caddy >/dev/null 2>&1; then
    # В репозитории дистрибутива нет — официальный репозиторий Caddy.
    apt-get install -y -qq debian-keyring debian-archive-keyring apt-transport-https gnupg >/dev/null
    curl -1sLf 'https://dl.cloudsmith.io/public/caddy/stable/gpg.key' \
      | gpg --dearmor -o /usr/share/keyrings/caddy-stable-archive-keyring.gpg
    curl -1sLf 'https://dl.cloudsmith.io/public/caddy/stable/debian.deb.txt' \
      > /etc/apt/sources.list.d/caddy-stable.list
    apt-get update -qq
    apt-get install -y -qq caddy >/dev/null
  fi
fi

cat > /etc/caddy/Caddyfile <<EOF
# Кабинет JobSignal. Создано deploy/expose_dashboard.sh.
$HOST {
	encode gzip
	header {
		Strict-Transport-Security "max-age=31536000"
		-Server
	}
	reverse_proxy 127.0.0.1:5000
}
EOF

# ── фаервол ──────────────────────────────────────────────────────────────────
# 80 нужен Let's Encrypt для проверки и для переадресации на https.
ufw allow 80/tcp >/dev/null
ufw allow 443/tcp >/dev/null
ufw deny 5000 >/dev/null 2>&1 || true

# Дашборд должен слушать только localhost: DASHBOARD_HOST наружу не нужен.
if grep -qE '^DASHBOARD_HOST=' "$ENV"; then
  sed -i 's|^DASHBOARD_HOST=.*|DASHBOARD_HOST=127.0.0.1|' "$ENV"
fi

systemctl restart jobsignal-dashboard
systemctl enable caddy >/dev/null 2>&1
systemctl restart caddy

# ── проверка ─────────────────────────────────────────────────────────────────
echo "==> жду сертификат (до минуты)"
code=""
for _ in $(seq 1 20); do
  code=$(curl -s -o /dev/null -w '%{http_code}' --max-time 10 "https://$HOST/status" || true)
  [ "$code" = "401" ] && break
  sleep 3
done

echo
if [ "$code" = "401" ]; then
  echo "ГОТОВО. Кабинет: https://$HOST/status"
else
  echo "Кабинет пока не ответил как ожидалось (код: ${code:-нет ответа})."
  echo "Посмотри: journalctl -u caddy -n 40 --no-pager"
fi
echo "Логин: $USER_NAME"
if [ -n "$NEWPASS" ]; then
  echo "Пароль: $NEWPASS"
  echo "(новый, записан в $ENV; сохрани его в менеджере паролей — больше он не покажется."
  echo " В чаты и переписку его не вставляй.)"
else
  echo "Пароль: прежний, из DASHBOARD_PASS в $ENV"
fi
