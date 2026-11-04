"""Chép đúng danh sách commit_paths.txt ra một thư mục trống ngoài repo, không dùng git.

Chạy từ gốc repo: python tools/commit_scan/export.py THU_MUC_TRONG [--only PHAN_DAU ...]
Dùng để đóng băng phần M1 trong lúc làm M2, và để dựng Docker, chạy test từ bản xuất sạch (Q77).
"""
from __future__ import annotations

import argparse
import hashlib
import json
import pathlib
import shutil
import sys

sys.path.insert(0, str(pathlib.Path(__file__).parent))

import scan  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("out", type=pathlib.Path)
    parser.add_argument("--only", nargs="*", default=None)
    args = parser.parse_args()
    out = args.out.resolve()
    if out == scan.ROOT or out.is_relative_to(scan.ROOT):
        raise SystemExit("Bản xuất phải nằm ngoài repo.")
    if out.exists() and any(out.iterdir()):
        raise SystemExit("Thư mục xuất phải trống.")
    files = [f for f in scan.commit_files() if args.only is None or any(f == p or f.startswith(p.rstrip("/") + "/") for p in args.only)]
    sums = {}
    for name in files:
        target = out / name
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(scan.ROOT / name, target)
        sums[name] = hashlib.sha256(target.read_bytes()).hexdigest()
    (out.parent / f"{out.name}.sha256.json").write_text(json.dumps(sums, indent=0), encoding="utf8")
    print(f"Đã xuất {len(files)} file vào {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
