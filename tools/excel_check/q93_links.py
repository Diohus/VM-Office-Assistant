"""Kiểm 6b.1 (Q93) trên Excel thật: biểu mẫu có liên kết ngoài, xuất từ app rồi mở trên máy không có file được liên kết.

Chạy khi app đang bật: python q93_links.py [http://127.0.0.1:8089]. Tải fixture X18b lên, sửa một ô không ô nào phụ thuộc, lưu, xuất.
Mở file gốc và file xuất trong thư mục tạm không có file nguồn, so mọi ô công thức ở từng cách mở. Ghi results/q93_links.json. Trả mã 1 nếu có ô khác.
"""
from __future__ import annotations

import json
import math
import pathlib
import sys
import urllib.request

from excel import excel

ROOT = pathlib.Path(__file__).resolve().parents[2]
FIXTURE = ROOT / "fixtures" / "templates" / "GIA-X18b_lien_ket_ngoai.xlsx"
OUT = pathlib.Path(__file__).parent / "results" / "q93_links.json"
SHEET = "KQ_QC_Thang"
LINK_CELLS = ["I8", "I9"]
TOL = 1e-9


def call(base: str, path: str, data: bytes | None = None, headers: dict | None = None) -> bytes:
    req = urllib.request.Request(base + path, data=data, headers=headers or {}, method="POST" if data is not None else "GET")
    with urllib.request.urlopen(req, timeout=120) as res:
        return res.read()


def export_from_app(base: str) -> tuple[bytes, dict]:
    doc = json.loads(call(base, "/api/documents?filename=GIA-X18b_lien_ket_ngoai.xlsx", FIXTURE.read_bytes(), {"Content-Type": "application/octet-stream"}))
    body = {"base_version": doc["version"], "author": "Người Giả", "edits": [{"sheet": SHEET, "cell": "K40", "value": 1}]}
    saved = json.loads(call(base, f"/api/documents/{doc['document_id']}/versions", json.dumps(body).encode("utf8"), {"Content-Type": "application/json"}))
    return call(base, f"/api/documents/{doc['document_id']}/export"), {"version": saved["version"], "external": saved.get("external")}


def read_formulas(wb) -> dict:
    out = {}
    for ws in wb.Worksheets:
        for row in ws.UsedRange.Rows:
            for c in row.Cells:
                if c.HasFormula:
                    out[f"{ws.Name}!{c.GetAddress(False, False)}"] = c.Value2
    return out


def same(a, b) -> bool:
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return math.isclose(a, b, rel_tol=TOL, abs_tol=TOL)
    return a == b


def read_one(xl, path: pathlib.Path, update_links: int, recalc: bool) -> dict:
    wb = xl.app.Workbooks.Open(str(xl.work.require_inside(path)), update_links)
    if recalc:
        xl.app.CalculateFull()
    values = read_formulas(wb)
    wb.Close(False)
    return values


def main() -> int:
    base = sys.argv[1] if len(sys.argv) > 1 else "http://127.0.0.1:8089"
    exported, meta = export_from_app(base)
    result: dict = {"app": meta, "modes": {}}
    bad = 0
    with excel() as xl:
        orig = xl.work.copy_in(FIXTURE, "goc.xlsx")
        out = xl.work.write("xuat.xlsx", exported)
        assert not (xl.work.path / "GIA-X18_nguon_ngoai.xlsx").exists(), "file được liên kết không được có mặt"
        for mode, (links, recalc) in {"khong_cap_nhat_lien_ket": (0, False), "dong_y_cap_nhat_lien_ket": (3, False), "tinh_lai_het": (0, True)}.items():
            a = read_one(xl, orig, links, recalc)
            b = read_one(xl, out, links, recalc)
            diffs = [{"cell": k, "goc": a.get(k), "xuat": b.get(k)} for k in sorted(set(a) | set(b)) if not same(a.get(k), b.get(k))]
            errors = [k for k, v in b.items() if isinstance(v, int) and v < -2146826000]
            result["modes"][mode] = {
                "formula_cells": len(b),
                "links_goc": {c: a.get(f"{SHEET}!{c}") for c in LINK_CELLS},
                "links_xuat": {c: b.get(f"{SHEET}!{c}") for c in LINK_CELLS},
                "diffs": diffs[:50],
                "diff_count": len(diffs),
                "error_cells_xuat": errors,
            }
            bad += len(diffs) + (0 if b else 1)
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text(json.dumps(result, ensure_ascii=False, indent=1), encoding="utf8")
    print(json.dumps(result, ensure_ascii=False, indent=1)[:3000])
    return 1 if bad else 0


if __name__ == "__main__":
    raise SystemExit(main())
