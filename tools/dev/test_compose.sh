#!/bin/sh
# docker compose cho RIÊNG bộ test (dự án vm-office-test, cổng 8089, máy giả fakeai, khóa mồi). Không bao giờ chạm app 8088.
# Dùng thay cho gõ tay: sh tools/dev/test_compose.sh up -d --no-build --force-recreate --wait web
# (PM 14/11: một lệnh gõ tay thiếu WEB_PORT=8089 đã suýt gắn bộ test vào cổng 8088.)
set -eu
cd "$(dirname "$0")/../.."
ROOT_DIR=$(pwd -W 2>/dev/null || pwd)
export VM_PROJECT=vm-office-test WEB_PORT=8089
export LLM_SECRET_FILE=tests/integration/out/llm_canary.env
export COMPOSE_PATH_SEPARATOR=";" COMPOSE_FILE="compose.yaml;tests/integration/compose.llm-test.yaml"
exec docker compose -p "$VM_PROJECT" --project-directory "$ROOT_DIR" --env-file .local/compose.env "$@"
