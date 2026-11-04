"""Bản chụp mã nguồn khi chưa được commit (Q35): file zip kèm danh sách SHA-256, để ngoài repo.

    python tools/snapshot/snapshot.py create [--dest THU_MUC]
    python tools/snapshot/snapshot.py diff A.zip B.zip
    python tools/snapshot/snapshot.py restore A.zip THU_MUC_TRONG

Không dùng git. Không chứa bí mật, môi trường cài đặt, file sinh ra hay biểu mẫu thật. Chỉ dùng thư viện chuẩn.
"""
from __future__ import annotations

import argparse
import datetime as dt
import fnmatch
import hashlib
import json
import pathlib
import sys
import zipfile

ROOT = pathlib.Path(__file__).resolve().parents[2]
DEFAULT_DEST = ROOT.parent / f"{ROOT.name}_snapshots"
MANIFEST = "SHA256SUMS.json"

EXCLUDE_DIRS = {".git", ".venv", "node_modules", "__pycache__", ".pytest_cache", "work", "out", ".tools", ".models", "dist", ".local",
                "generated", "real", "real_forms", "bieu_mau_that"}
EXCLUDE_FILES = (".env", ".env.*", "~$*", "*.pem", "*.key", "*.pyc", "*.egg-info", "machine-ca*")
ALLOW_FILES = (".env.example",)
CLOUD_MARKERS = ("onedrive", "dropbox", "google drive", "googledrive", "icloud")


def _excluded(rel: pathlib.PurePath) -> bool:
    if set(rel.parts[:-1]) & EXCLUDE_DIRS or any(p.endswith(".egg-info") for p in rel.parts):
        return True
    name = rel.name
    return name not in ALLOW_FILES and any(fnmatch.fnmatch(name, pat) for pat in EXCLUDE_FILES)


def _check_dest(dest: pathlib.Path) -> None:
    resolved = dest.resolve()
    if resolved == ROOT or resolved.is_relative_to(ROOT):
        raise SystemExit("Bản chụp phải nằm ngoài repo.")
    lowered = str(resolved).lower()
    if any(marker in lowered for marker in CLOUD_MARKERS):
        raise SystemExit("Không để bản chụp trong thư mục đồng bộ cloud.")


def create(dest: pathlib.Path) -> pathlib.Path:
    _check_dest(dest)
    dest.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    target = dest / f"snapshot-{stamp}.zip"
    sums: dict[str, str] = {}
    skipped: list[str] = []
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in sorted(ROOT.rglob("*")):
            rel = path.relative_to(ROOT)
            if not path.is_file() or _excluded(rel):
                continue
            try:
                data = path.read_bytes()
            except OSError:      # file đang bị tiến trình khác giữ khóa: báo, không bỏ qua trong im lặng
                skipped.append(rel.as_posix())
                continue
            sums[rel.as_posix()] = hashlib.sha256(data).hexdigest()
            zf.writestr(rel.as_posix(), data)
        zf.writestr(MANIFEST, json.dumps({"created": stamp, "files": sums, "skipped": skipped}, ensure_ascii=False, indent=0))
    print(f"{target}  ({len(sums)} file, {target.stat().st_size / 1024 / 1024:.1f} MB)")
    for name in skipped:
        print(f"Không đọc được, không có trong bản chụp: {name}")
    return target


def _manifest(path: pathlib.Path) -> dict:
    with zipfile.ZipFile(path) as zf:
        return json.loads(zf.read(MANIFEST))


def diff(a: pathlib.Path, b: pathlib.Path) -> dict:
    fa, fb = _manifest(a)["files"], _manifest(b)["files"]
    res = {"them": sorted(set(fb) - set(fa)), "xoa": sorted(set(fa) - set(fb)), "doi": sorted(k for k in set(fa) & set(fb) if fa[k] != fb[k])}
    for key, names in res.items():
        for name in names:
            print(f"{key}: {name}")
    print(f"Thêm {len(res['them'])}, xóa {len(res['xoa'])}, đổi {len(res['doi'])}.")
    return res


def restore(source: pathlib.Path, target: pathlib.Path) -> int:
    if target.exists() and any(target.iterdir()):
        raise SystemExit("Thư mục khôi phục phải trống, để không ghi đè gì.")
    manifest = _manifest(source)
    root = target.resolve()
    with zipfile.ZipFile(source) as zf:
        for name, expected in manifest["files"].items():
            out = (root / name).resolve()
            if not out.is_relative_to(root):
                raise SystemExit(f"Đường dẫn không hợp lệ trong bản chụp: {name}")
            data = zf.read(name)
            if hashlib.sha256(data).hexdigest() != expected:
                raise SystemExit(f"Sai hash, bản chụp hỏng: {name}")
            out.parent.mkdir(parents=True, exist_ok=True)
            out.write_bytes(data)
    print(f"Đã khôi phục {len(manifest['files'])} file vào {root}, hash khớp hết.")
    return len(manifest["files"])


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="cmd", required=True)
    c = sub.add_parser("create")
    c.add_argument("--dest", type=pathlib.Path, default=DEFAULT_DEST)
    d = sub.add_parser("diff")
    d.add_argument("a", type=pathlib.Path)
    d.add_argument("b", type=pathlib.Path)
    r = sub.add_parser("restore")
    r.add_argument("source", type=pathlib.Path)
    r.add_argument("target", type=pathlib.Path)
    args = parser.parse_args()
    if args.cmd == "create":
        create(args.dest)
    elif args.cmd == "diff":
        diff(args.a, args.b)
    else:
        restore(args.source, args.target)
    return 0


if __name__ == "__main__":
    sys.exit(main())
