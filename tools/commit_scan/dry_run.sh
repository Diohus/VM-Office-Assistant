#!/bin/sh
# Thử trước commit đầu tiên, không đụng git: chép đúng danh sách commit ra thư mục tạm, quét, rồi chạy test reader ở đó.
# Mục đích: chứng minh bản commit tự đứng được, không có tài liệu nội bộ, .env. Chạy từ gốc repo: sh tools/commit_scan/dry_run.sh
# Tách commit (Q127): DRY_RUN_KEEP_HEAD="vmai/ fixtures/eval/" sh tools/commit_scan/dry_run.sh
#   các đường dẫn đó không lấy từ cây làm việc mà lấy bản ở HEAD, đúng như kho sau commit A.
set -eu
ROOT="$(pwd)"
PY="$ROOT/workers/reader/.venv/Scripts/python.exe"
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
import os, subprocess, tarfile, io
keep = os.environ.get("DRY_RUN_KEEP_HEAD", "").split()
files = [f for f in files if not any(f.startswith(k) for k in keep)]
for f in files:
    target = out / f
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(f, target)
if keep:
    data = subprocess.run(["git", "archive", "HEAD", *keep], capture_output=True, check=True).stdout
    with tarfile.open(fileobj=io.BytesIO(data)) as t:
        t.extractall(out, filter="data")
    print("Lấy từ HEAD:", " ".join(keep))
size = sum((out / f).stat().st_size for f in files)
print(f"Danh sách commit: {len(files)} file, {size / 1024 / 1024:.1f} MB")
for f in sorted(files, key=lambda f: -(out / f).stat().st_size)[:5]:
    print(" ", f, round((out / f).stat().st_size / 1024), "KB")
PY
cd "$OUT"
SEP=":"; case "$(uname -s)" in MINGW*|MSYS*) SEP=";" ;; esac
PYTHONPATH="workers/reader${SEP}vmai/src" "$PY" -m pytest workers/reader/tests -q -W ignore -p no:cacheprovider
# Q128: test của vmai trên bản chép, bằng môi trường Python của vmai trong repo. In số đạt, bỏ qua, lỗi; có lỗi thì dừng.
VPY="$ROOT/vmai/.venv/Scripts/python.exe"
[ -x "$VPY" ] || VPY="$ROOT/vmai/.venv/bin/python"
VOUT="$(mktemp)"
if (cd vmai && PYTHONPATH=src "$VPY" -m pytest -q -p no:cacheprovider tests > "$VOUT" 2>&1); then
  echo "vmai: $(tail -1 "$VOUT")"
else
  tail -20 "$VOUT"; rm -f "$VOUT"; echo "vmai: có test hỏng"; exit 1
fi
rm -f "$VOUT"
test -e workers/reader/reader/writer.py && test -e workers/libreoffice/worker.py && test -e frontend/src/numberFormat.js
test ! -e docs && test ! -e spikes && test ! -e chat_log.md && test ! -e .env
echo "Đạt: có đủ phần M2, không có tài liệu nội bộ, .env."
