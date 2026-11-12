#!/bin/sh
# Tầng 1: bật app, test reader không cần Docker, rồi test API và trình duyệt. Chạy từ gốc repo: sh tests/integration/run.sh
set -eu
PY="${VM_PY:-workers/reader/.venv/Scripts/python.exe}"
[ -x "$PY" ] || PY=workers/reader/.venv/bin/python
SEP=":"; case "$(uname -s)" in MINGW*|MSYS*) SEP=";" ;; esac
[ -f fixtures/generated/GIA-CONG_THUC_300x25.xlsx ] || "$PY" tools/fixtures/make_limit_files.py > /dev/null
# Bộ container riêng, cổng riêng: test có tắt API, không được đụng tới app đang có người dùng.
export VM_PROJECT=vm-office-test WEB_PORT=8089 VM_BASE=http://127.0.0.1:8089
mkdir -p tests/integration/out
# Lát 2 (Q135): llm gọi máy giả, key là khóa mồi sinh một lần. Không bao giờ gắn .env thật vào bộ test.
CANARY=tests/integration/out/llm_canary.env
[ -s "$CANARY" ] || "$PY" -c 'import secrets; print("OPENAI_API_KEY=sk-test-canary-" + secrets.token_hex(16))' > "$CANARY"
export LLM_SECRET_FILE="$CANARY" COMPOSE_PATH_SEPARATOR=";" COMPOSE_FILE="compose.yaml;tests/integration/compose.llm-test.yaml"
sh tools/dev/up.sh > tests/integration/out/up.log 2>&1 || { echo "Không bật được app:"; tail -25 tests/integration/out/up.log; exit 1; }
# fakeai chạy thẳng script gắn vào: compose không tạo lại container khi chỉ script đổi, nên khởi động lại để dùng bản mới.
docker restart "$VM_PROJECT-fakeai-1" > /dev/null
PYTHONIOENCODING=utf8 PYTHONPATH="workers/reader${SEP}vmai/src" "$PY" -m pytest workers/reader/tests -q -W ignore -p no:cacheprovider
PYTHONIOENCODING=utf8 PYTHONDONTWRITEBYTECODE=1 PYTHONPATH="workers/llm${SEP}workers/egress${SEP}backend" "$PY" -m pytest workers/llm/tests workers/egress/tests backend/tests -q -W ignore -p no:cacheprovider
[ -d tests/integration/browser/node_modules ] || (cd tests/integration/browser && npm ci --ignore-scripts --no-audit --no-fund > /dev/null)
(cd tests/integration/browser && node --test typed_number.test.mjs fill_model.test.mjs > /dev/null)
PYTHONIOENCODING=utf8 PYTHONDONTWRITEBYTECODE=1 "$PY" -m pytest tests/integration -q -W ignore -p no:cacheprovider "$@"
