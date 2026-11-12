"""Tính lại file giả của khung chat bằng LibreOffice, để eval tầng model thấy kết quả công thức thật (Q135, mốc C).

File do openpyxl lưu không có kết quả tính sẵn, nên model sẽ thấy "→ 0" ở mọi ô công thức. Script này chạy
một container MỚI từ image `vm-office-converter` (không động vào container đang chạy của Dev), tắt mạng,
gốc chỉ đọc, chỉ gắn thư mục `fixtures/templates/chat_lat2/`. Cài đặt tính giống worker PDF: locale en-GB, luôn tính lại.

Chạy từ gốc repo:  vmai/.venv/Scripts/python.exe fixtures/eval/recalc_chat_fixtures.py
Kết quả: `fixtures/templates/chat_lat2/tinh_lai/<tên>.xlsx`. Chỉ dùng để eval đọc kết quả, không phải biểu mẫu để sửa.
"""
from __future__ import annotations

import io
import re
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
DIR = ROOT / "fixtures" / "templates" / "chat_lat2"
OUT = DIR / "tinh_lai"
IMAGE = "vm-office-converter:latest"
FILES = ["GIA-QC-01_chat.xlsx"]

# Luôn tính lại khi mở (OOXMLRecalcMode 0), tính theo dấu chấm thập phân (en-GB), như worker PDF của Dev.
_SETTINGS = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Setup/L10N"><prop oor:name="ooSetupSystemLocale" oor:op="fuse"><value>en-GB</value></prop></item>
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>
"""
_SCRIPT = (
    "set -e; mkdir -p /tmp/p/user /tmp/out; cat > /tmp/p/user/registrymodifications.xcu; "
    "soffice -env:UserInstallation=file:" "//" "/tmp/p --headless --norestore "
    "--convert-to 'xlsx:Calc MS Excel 2007 XML' --outdir /tmp/out /in/{name} >/dev/null 2>&1; "
    "cat /tmp/out/{name}"
)
# LibreOffice ghi tên vùng rỗng làm Excel không mở được (worker.py, đo ở M0). Thông tin file: bỏ tác giả.
_EMPTY_NAME = re.compile(rb"<definedName\b[^>]*>\s*</definedName>|<definedName\b[^>]*/>")
_EMPTY_NAMES = re.compile(rb"<definedNames>\s*</definedNames>")
_PEOPLE = re.compile(rb"<(dc:creator|cp:lastModifiedBy)>[^<]*</\1>")


def _clean(data: bytes) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as zi, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for name in zi.namelist():
            part = zi.read(name)
            if name == "xl/workbook.xml":
                part = _EMPTY_NAMES.sub(b"", _EMPTY_NAME.sub(b"", part))
            if name == "docProps/core.xml":
                part = _PEOPLE.sub(b"", part)
            zo.writestr(name, part)
    return out.getvalue()


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    OUT.mkdir(exist_ok=True)
    for name in FILES:
        result = subprocess.run(
            ["docker", "run", "--rm", "-i", "--network", "none", "--read-only", "--tmpfs", "/tmp", "--cap-drop", "ALL",
             "--security-opt", "no-new-privileges", "-v", f"{DIR}:/in:ro", "--entrypoint", "sh", IMAGE,
             "-c", _SCRIPT.format(name=name)],
            input=_SETTINGS.encode(), capture_output=True, timeout=180,
        )
        if result.returncode != 0 or not result.stdout.startswith(b"PK"):
            print(f"Không tính lại được {name} (mã {result.returncode})")
            return 1
        (OUT / name).write_bytes(_clean(result.stdout))
        print(f"Đã tính lại {name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
