"""Rà lỗ hổng cho lockfile sản phẩm. Dùng kho chứng chỉ của Windows qua truststore, không tắt xác minh TLS.

Chạy (từ gốc repo, môi trường có pip-audit và truststore): python tools/deps/audit.py
Kết quả ghi vào docs/engineering/audit/. Trả mã khác 0 nếu có lỗ hổng hoặc không rà được.
"""
from __future__ import annotations

import json
import pathlib
import subprocess
import sys

import truststore

ROOT = pathlib.Path(__file__).resolve().parents[2]
OUT = ROOT / "docs" / "engineering" / "audit"
LOCKS = {"backend": ROOT / "backend/requirements.lock", "reader": ROOT / "workers/reader/requirements.lock"}
NPM = {"frontend": ROOT / "frontend", "tests": ROOT / "tests/integration/browser"}


def pip_audit(name: str, lock: pathlib.Path) -> int:
    truststore.inject_into_ssl()
    from pip_audit._cli import audit

    out = OUT / f"pip_audit_{name}.json"
    argv = sys.argv
    sys.argv = ["pip-audit", "-r", str(lock), "--require-hashes", "--disable-pip", "--format", "json", "--output", str(out), "--progress-spinner", "off"]
    try:
        audit()
    except SystemExit as exc:
        return int(exc.code or 0)
    finally:
        sys.argv = argv
    return 0


def npm_audit(name: str, folder: pathlib.Path) -> int:
    npm = "npm.cmd" if sys.platform == "win32" else "npm"
    proc = subprocess.run([npm, "audit", "--json"], cwd=folder, capture_output=True, text=True, encoding="utf8")
    (OUT / f"npm_audit_{name}.json").write_text(proc.stdout, encoding="utf8")
    data = json.loads(proc.stdout)
    print(name, data.get("metadata", {}).get("vulnerabilities"))
    return proc.returncode


def main() -> int:
    OUT.mkdir(parents=True, exist_ok=True)
    failed = 0
    for name, lock in LOCKS.items():
        rc = pip_audit(name, lock)
        report = json.loads((OUT / f"pip_audit_{name}.json").read_text(encoding="utf8"))
        vulns = [(d["name"], d["version"], v["id"]) for d in report["dependencies"] for v in d.get("vulns", [])]
        print(name, "gói:", len(report["dependencies"]), "lỗ hổng:", vulns, "rc:", rc)
        failed |= rc
    for name, folder in NPM.items():
        failed |= npm_audit(name, folder)
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
