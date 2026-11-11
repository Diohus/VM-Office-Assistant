"""Quét những gì sẽ vào commit: đường dẫn máy, tên người dùng, khóa. Xem cả bên trong zip của file Office và chuỗi trong file đời cũ.

Chạy từ gốc repo: python tools/commit_scan/scan.py        (quét đúng danh sách commit_paths.txt)
Trả mã 1 nếu có phát hiện. Chỉ in vị trí và loại, không in nội dung khớp, để không làm lộ khóa ra log.
"""
from __future__ import annotations

import fnmatch
import getpass
import io
import os
import pathlib
import re
import subprocess
import sys
import zipfile

ROOT = pathlib.Path(__file__).resolve().parents[2]
ZIP_SUFFIXES = {".xlsx", ".xlsm", ".docx", ".pptx", ".zip"}
_GENERATED_DIRS = {".local", ".venv", "node_modules", "__pycache__", ".pytest_cache", "out", "generated", "dist"}
# Các file này chứa mẫu để quét hoặc để lọc, không phải đường dẫn thật. Thêm vào đây phải đọc từng dòng trước.
TEXT_ALLOW = {"tools/commit_scan/scan.py", "tests/integration/test_commit_scan.py", "tools/excel_check/make_external_link.py", "vmai/tests/test_core.py"}


def _word_terms() -> list[bytes]:
    names = {getpass.getuser(), os.environ.get("USERNAME", ""), pathlib.Path.home().name}
    return [n.encode() for n in names if len(n) >= 4]


_RULES: list[tuple[str, re.Pattern[bytes]]] = [
    ("duong_dan_o_dia", re.compile(rb"(?<![A-Za-z0-9])[A-Za-z]:\\+[A-Za-z0-9_.$ -]")),
    ("duong_dan_file", re.compile(rb"file:///", re.I)),
    ("duong_dan_mang", re.compile(rb"\\\\[A-Za-z0-9_.-]+\\[A-Za-z0-9_$.-]+")),
    # /home/lo, /home/vm là người dùng trong container, do chính đội đặt.
    ("thu_muc_nguoi_dung", re.compile(rb"[\\/](?:Users|home)[\\/](?!(?:lo|vm)\b)[A-Za-z0-9_.-]+", re.I)),
    ("khoa_openai_anthropic", re.compile(rb"\bsk-[A-Za-z0-9_-]{20,}")),
    ("khoa_openrouter", re.compile(rb"\bsk-or-[A-Za-z0-9_-]{20,}")),
    ("khoa_aws", re.compile(rb"\bAKIA[0-9A-Z]{16}\b")),
    ("khoa_github", re.compile(rb"\bgh[pousr]_[A-Za-z0-9]{30,}")),
    ("khoa_rieng", re.compile(rb"-----BEGIN [A-Z ]*PRIVATE KEY-----")),
    ("gan_khoa", re.compile(rb"""(?i)\b(?:api[_-]?key|secret|password|token)\b\s*[:=]\s*['"][A-Za-z0-9_\-/+]{16,}['"]""")),
]


def _variants(data: bytes) -> list[bytes]:
    """File đời cũ lưu chuỗi dạng UTF-16: bỏ byte 0 xen kẽ để quét được."""
    return [data, data.replace(b"\x00", b"")] if b"\x00" in data else [data]


def scan_bytes(data: bytes, extra_terms: list[bytes] | None = None) -> list[str]:
    found = []
    terms = (extra_terms if extra_terms is not None else _word_terms())
    for view in _variants(data):
        for name, rule in _RULES:
            if rule.search(view):
                found.append(name)
        # Nguyên từ: tên máy "nguye" không được khớp nhầm "nguyen_ban".
        if any(re.search(rb"(?<![A-Za-z0-9])" + re.escape(t) + rb"(?![A-Za-z0-9])", view, re.I) for t in terms):
            found.append("ten_nguoi_dung")
    return sorted(set(found))


def scan_file(data: bytes, name: str, extra_terms: list[bytes] | None = None) -> list[tuple[str, str]]:
    """Trả [(phần trong file, loại phát hiện)]."""
    out: list[tuple[str, str]] = []
    if pathlib.PurePath(name).suffix.lower() in ZIP_SUFFIXES and data[:4] == b"PK\x03\x04":
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as zf:
                for part in zf.namelist():
                    out += [(part, kind) for kind in scan_bytes(zf.read(part), extra_terms)]
            return out
        except zipfile.BadZipFile:
            pass  # file cố tình hỏng: quét thẳng byte
    return [("", kind) for kind in scan_bytes(data, extra_terms)]


def commit_files() -> list[str]:
    """Các file sẽ vào commit: theo danh sách, trừ phần git bỏ qua."""
    lines = (ROOT / "tools/commit_scan/commit_paths.txt").read_text(encoding="utf8").splitlines()
    include = [ln.strip() for ln in lines if ln.strip() and not ln.startswith(("#", "!"))]
    exclude = [ln[1:].strip() for ln in lines if ln.startswith("!")]
    if (ROOT / ".git").exists():
        listed = subprocess.run(["git", "ls-files", "-z", "-c", "-o", "--exclude-standard"], cwd=ROOT, capture_output=True, check=True).stdout.decode("utf8").split("\0")
    else:
        # Bản xuất sạch không phải repo git: liệt kê thẳng, chỉ bỏ những gì do chạy test sinh ra.
        listed = [p.relative_to(ROOT).as_posix() for p in ROOT.rglob("*") if p.is_file() and not set(p.relative_to(ROOT).parts) & _GENERATED_DIRS]
    out = []
    # File đã theo dõi nhưng đã xóa hoặc chuyển chỗ trong cây làm việc: commit sẽ xóa nó, không có gì để quét.
    for path in sorted(p for p in listed if p and (ROOT / p).is_file()):
        if any(path == i or path.startswith(i.rstrip("/") + "/") for i in include) and not any(fnmatch.fnmatch(path, e) or path.startswith(e.rstrip("/") + "/") for e in exclude):
            out.append(path)
    return out


def main() -> int:
    files = commit_files()
    bad = 0
    for path in files:
        if path in TEXT_ALLOW:
            continue
        for part, kind in scan_file((ROOT / path).read_bytes(), path):
            print(f"{path}{' :: ' + part if part else ''}: {kind}")
            bad += 1
    print(f"Đã quét {len(files)} file, {bad} phát hiện.")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
