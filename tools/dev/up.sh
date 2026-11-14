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
  # Lối tắt Desktop (mo_app.cmd) chỉ đọc .local/compose.env. Ghi lại đường dẫn vừa chọn (không phải nội dung) để hai đường mở app
  # dùng cùng một file key. Chỉ ghi cho app thật và khi chính script này chọn: bộ test (VM_PROJECT khác, khóa mồi do người gọi
  # đặt) không bao giờ ghi, vì compose.env dùng chung với app thật.
  if [ "${VM_PROJECT:-vm-office}" = "vm-office" ]; then
    umask 077
    grep -v '^LLM_SECRET_FILE=' .local/compose.env > .local/compose.env.tmp || :
    printf 'LLM_SECRET_FILE=%s\n' "$LLM_SECRET_FILE" >> .local/compose.env.tmp
    mv .local/compose.env.tmp .local/compose.env
  fi
fi
export LLM_SECRET_FILE
# R-1020: từ Docker 29.8 (Compose v5, buildx bake), thư mục dự án tương đối làm bake ghi Dockerfile thành "workers\egress\Dockerfile"
# rồi đòi quyền --allow=fs.read, nên không dựng được image. Đưa thư mục dự án tuyệt đối (dạng D:/... trên Git Bash) thì bake
# thấy Dockerfile nằm trong ngữ cảnh dựng, không cần thêm quyền. Không tắt kiểm quyền của bake.
ROOT_DIR=$(pwd -W 2>/dev/null || pwd)
docker compose -p "${VM_PROJECT:-vm-office}" --project-directory "$ROOT_DIR" --env-file .local/compose.env up -d --build --wait "$@"
echo "App: http://127.0.0.1:${WEB_PORT:-8088}"
