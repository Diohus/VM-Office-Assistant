#!/bin/sh
# Bật toàn bộ app trên máy này. Chạy từ gốc repo: sh tools/dev/up.sh
# Mật khẩu Postgres tự sinh, nằm ở .local/compose.env. File .env ở gốc repo dành riêng cho khóa AI của khách hàng, script này không đụng tới.
set -eu
if [ -f .local/compose.env ]; then
  # Bản cũ của script có thể đã ghi mật khẩu rỗng khi máy thiếu Python. Không tự xóa gì, chỉ báo.
  if ! grep -q '^POSTGRES_PASSWORD=..*' .local/compose.env; then
    echo "Chưa bật được app: file .local/compose.env có mật khẩu kho dữ liệu bị trống, có thể do lần cài trước máy chưa có Python."
    echo "Mật khẩu trống thì kho dữ liệu không tạo được, nên chưa có bản nào được lưu bằng mật khẩu này."
    echo "Hãy kiểm tra Python (gõ: python --version), rồi xóa riêng file .local/compose.env và chạy lại lệnh này."
    exit 1
  fi
else
  PASSWORD=""
  for py in python python3; do
    PASSWORD=$("$py" -c 'import secrets; print(secrets.token_urlsafe(24))' 2>/dev/null) && [ ${#PASSWORD} -ge 20 ] && break
    PASSWORD=""
  done
  if [ -z "$PASSWORD" ]; then
    echo "Chưa cài được app: máy này chưa có Python 3, hoặc gõ python trong Git Bash chưa chạy được Python thật."
    echo "Chưa có gì được tạo hay thay đổi. Hãy cài Python 3 từ python.org, chọn \"Add python.exe to PATH\", mở lại Git Bash rồi chạy lại lệnh này."
    exit 1
  fi
  mkdir -p .local
  umask 077
  printf 'POSTGRES_PASSWORD=%s\n' "$PASSWORD" > .local/compose.env
fi
if [ ! -f .local/machine-ca.pem ]; then
  powershell -NoProfile -ExecutionPolicy Bypass -File tools/deps/export_machine_ca.ps1 2>/dev/null || : > .local/machine-ca.pem
fi
# Key AI (Q135): chỉ chọn đường dẫn file, không đọc nội dung. Docker gắn file vào riêng dịch vụ llm.
# Bộ test đặt LLM_SECRET_FILE là file khóa mồi, không bao giờ dùng .env thật. Chưa có .env thì llm chạy không có key.
: > .local/llm_none.env
if [ -z "${LLM_SECRET_FILE:-}" ]; then
  if [ -f .env ]; then LLM_SECRET_FILE=.env; else LLM_SECRET_FILE=.local/llm_none.env; fi
fi
export LLM_SECRET_FILE
docker compose -p "${VM_PROJECT:-vm-office}" --env-file .local/compose.env up -d --build --wait "$@"
echo "App: http://127.0.0.1:${WEB_PORT:-8088}"
