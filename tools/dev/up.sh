#!/bin/sh
# Bật toàn bộ app trên máy này. Chạy từ gốc repo: sh tools/dev/up.sh
# Mật khẩu Postgres tự sinh, nằm ở .local/compose.env. File .env ở gốc repo dành riêng cho khóa AI của khách hàng, script này không đụng tới.
set -eu
mkdir -p .local
if [ ! -f .local/compose.env ]; then
  umask 077
  printf 'POSTGRES_PASSWORD=%s\n' "$(python -c 'import secrets; print(secrets.token_urlsafe(24))')" > .local/compose.env
fi
if [ ! -f .local/machine-ca.pem ]; then
  powershell -NoProfile -ExecutionPolicy Bypass -File tools/deps/export_machine_ca.ps1 2>/dev/null || : > .local/machine-ca.pem
fi
docker compose -p "${VM_PROJECT:-vm-office}" --env-file .local/compose.env up -d --build --wait "$@"
echo "App: http://127.0.0.1:${WEB_PORT:-8088}"
