#!/bin/sh
# Thử trước commit đầu tiên, không đụng git: chép đúng danh sách commit ra thư mục tạm, quét, rồi chạy test reader ở đó.
# Mục đích: chứng minh phần M1 tự đứng được khi không có phần M2 đang làm. Chạy từ gốc repo: sh tools/commit_scan/dry_run.sh
set -eu
PY="$(pwd)/workers/reader/.venv/Scripts/python.exe"
[ -x "$PY" ] || PY="$(pwd)/workers/reader/.venv/bin/python"
export PYTHONIOENCODING=utf8
"$PY" tools/commit_scan/scan.py
OUT="$(mktemp -d)"
trap 'rm -rf "$OUT"' EXIT
"$PY" - "$OUT" <<'PY'
import pathlib, shutil, sys
sys.path.insert(0, "tools/commit_scan")
import scan
out = pathlib.Path(sys.argv[1])
files = scan.commit_files()
for f in files:
    target = out / f
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(f, target)
size = sum((out / f).stat().st_size for f in files)
print(f"Danh sách commit: {len(files)} file, {size / 1024 / 1024:.1f} MB")
for f in sorted(files, key=lambda f: -(out / f).stat().st_size)[:5]:
    print(" ", f, round((out / f).stat().st_size / 1024), "KB")
PY
cd "$OUT"
SEP=":"; case "$(uname -s)" in MINGW*|MSYS*) SEP=";" ;; esac
PYTHONPATH="workers/reader${SEP}vmai/src" "$PY" -m pytest workers/reader/tests -q -W ignore -p no:cacheprovider
test ! -e workers/reader/reader/writer.py
test ! -e docs && test ! -e spikes && test ! -e chat_log.md && test ! -e .env
echo "Đạt: không có phần M2, tài liệu nội bộ, .env."
