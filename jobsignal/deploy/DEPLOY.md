# Деплой на сервер (Ubuntu)

Цель: круглосуточный автосбор + матчинг + всегда доступный дашборд, без твоего компьютера.
Отправка рекрутёрам остаётся ручной (semi-auto) — сервер делает всё до неё.

## 0. Переустановить ОС
В панели Timeweb: сервер → «Переустановить ОС» → **Ubuntu 24.04 LTS**. Получишь чистую машину.

## 1. Залить проект на сервер
С локального компьютера (Mac), из папки, где лежит распакованный проект:
```bash
scp -r jobsignal_local root@5.42.104.39:/opt/jobsignal_local
```
(пароль root — из панели Timeweb; либо настрой SSH-ключ заранее)

## 2. Запустить установку
Подключиться и выполнить скрипт — он сделает swap, зависимости, автозапуск и фаервол:
```bash
ssh root@5.42.104.39
cd /opt/jobsignal_local
bash deploy/setup_server.sh
```

## 3. Вписать ключи
Скрипт создаст `config/.env` из шаблона. Впиши свои значения:
```bash
nano config/.env
```
Обязательно:
- `GIGACHAT_AUTH_KEY=...` — ключ авторизации (перевыпусти новый!)
- `DASHBOARD_PASS=...` — пароль на дашборд (на публичном IP без него нельзя)
- `LLM_PROVIDER=gigachat`

Перенести каналы и резюме можно тем же `scp` (папка `config/`), или они уже уехали с проектом.

После правки — перезапустить:
```bash
systemctl restart jobsignal-dashboard
```

## 4. Первый сбор и проверка
```bash
systemctl start jobsignal-pipeline.service     # разовый прогон сейчас
journalctl -u jobsignal-pipeline -f            # смотреть лог сбора/матчинга
```
Дашборд: `http://5.42.104.39:5000` (логин/пароль из .env).

## Эксплуатация
- Конвейер идёт сам каждые 6 часов (таймер systemd).
- Логи: `journalctl -u jobsignal-dashboard -f` и `journalctl -u jobsignal-pipeline -f`.
- Изменить частоту: правь `OnUnitActiveSec` в `/etc/systemd/system/jobsignal-pipeline.timer`, затем `systemctl daemon-reload && systemctl restart jobsignal-pipeline.timer`.
- Обновить код: повторный `scp` поверх + `systemctl restart jobsignal-dashboard`.

## Память (1 ГБ)
Swap 2 ГБ ставится автоматически. LLM считается в облаке GigaChat, локально модель не грузится,
поэтому 1 ГБ + swap для сбора/матчинга/дашборда достаточно.
