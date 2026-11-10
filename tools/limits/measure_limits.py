"""Đo các ca giới hạn nhẹ qua API trên bộ thử 8089 (T-DEV-05). Không đụng 8088, chỉ dùng file giả.

Chạy từ gốc repo, sau khi bộ 8089 đã bật (sh tests/integration/run.sh hoặc VM_PROJECT=vm-office-test WEB_PORT=8089 sh tools/dev/up.sh):
    workers/reader/.venv/Scripts/python.exe tools/limits/measure_limits.py
Kết quả: tests/integration/out/limits.json. Ca nặng (file lớn hợp lệ, xuất PDF đồng thời) không nằm ở đây.
"""
from __future__ import annotations

import io
import json
import os
import pathlib
import subprocess
import time
import urllib.error
import urllib.parse
import urllib.request
import zipfile

ROOT = pathlib.Path(__file__).resolve().parents[2]
FIX = ROOT / "fixtures" / "templates"
GEN = ROOT / "fixtures" / "generated"
OUT = ROOT / "tests" / "integration" / "out"
BASE = os.environ.get("VM_BASE", "http://127.0.0.1:8089")
PROJECT = os.environ.get("VM_PROJECT", "vm-office-test")
SHEET = "KQ_QC_Thang"
QC = FIX / "GIA-QC-01_mau.xlsx"

if "8088" in BASE or PROJECT == "vm-office":
    raise SystemExit("Chỉ chạy trên bộ thử 8089, không chạy trên app đang nghiệm thu.")


def call(method: str, path: str, data: bytes | None = None, headers: dict | None = None, timeout: int = 180):
    req = urllib.request.Request(BASE + path, data=data, method=method, headers=headers or {})
    started = time.monotonic()
    try:
        with urllib.request.urlopen(req, timeout=timeout) as res:
            return res.status, dict(res.headers), res.read(), round(time.monotonic() - started, 2)
    except urllib.error.HTTPError as err:
        return err.code, dict(err.headers), err.read(), round(time.monotonic() - started, 2)


def body_of(raw: bytes):
    try:
        return json.loads(raw)
    except ValueError:
        return {"raw": raw[:200].decode("utf8", "replace")}


def upload(data: bytes, name: str) -> dict:
    status, _, raw, seconds = call("POST", f"/api/documents?filename={urllib.parse.quote(name)}", data)
    body = body_of(raw)
    return {"status": status, "seconds": seconds, "code": body.get("code"), "message": body.get("message"), "document_id": body.get("document_id"), "bytes": len(data)}


def post_json(path: str, payload: dict):
    status, headers, raw, seconds = call("POST", path, json.dumps(payload).encode(), {"Content-Type": "application/json"})
    return status, headers, raw, seconds


def store_bytes() -> int | None:
    try:
        out = subprocess.run(["docker", "exec", f"{PROJECT}-api-1", "du", "-sb", "/store"], capture_output=True, text=True, timeout=60)
        return int(out.stdout.split()[0])
    except (OSError, ValueError, IndexError, subprocess.TimeoutExpired):
        return None


def padded(size_bytes: int) -> bytes:
    """Biểu mẫu giả cộng một phần lưu nguyên, không nén, cho đủ cỡ."""
    out = io.BytesIO()
    with zipfile.ZipFile(QC) as zi, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for n in zi.namelist():
            zo.writestr(n, zi.read(n))
        extra = size_bytes - QC.stat().st_size
        zo.writestr(zipfile.ZipInfo("customXml/dem.bin"), os.urandom(max(extra, 0)), compress_type=zipfile.ZIP_STORED)
    return out.getvalue()


def main() -> None:
    res: dict = {"base": BASE, "started": time.strftime("%Y-%m-%d %H:%M:%S")}
    qc = QC.read_bytes()

    # GH-05: cỡ file, số ô, giải nén.
    res["GH-05b_25.5MiB"] = upload((GEN / "GIA-VUOT_25MB.xlsx").read_bytes(), "GIA-VUOT_25MB.xlsx")
    res["GH-05c_27MiB"] = upload(padded(27 * 1024 * 1024), "GIA-VUOT_27MB.xlsx")
    res["GH-05_24MiB_hop_le"] = upload(padded(24 * 1024 * 1024), "GIA-24MB.xlsx")
    res["GH-05d_200000_o"] = upload((GEN / "GIA-VUOT_200000_o.xlsx").read_bytes(), "GIA-VUOT_200000_o.xlsx")
    res["GH-05e_giai_nen"] = upload((GEN / "GIA-VUOT_giai_nen_200MB.xlsx").read_bytes(), "GIA-VUOT_giai_nen_200MB.xlsx")

    # GH-07: file lạ.
    res["GH-07a_doi_duoi"] = upload("đây là chữ, không phải Excel\n".encode("utf8") * 50, "ghi_chu.xlsx")
    res["GH-07b_hong"] = upload(qc[: len(qc) // 2], "GIA-QC-01_chep_thieu.xlsx")
    res["GH-07c_xls_mat_khau"] = upload((FIX / "GIA-X19_xls_mat_khau.xls").read_bytes(), "GIA-X19_xls_mat_khau.xls")
    res["GH-07c_xlsx_mat_khau"] = upload((FIX / "GIA-X19_co_mat_khau.xlsx").read_bytes(), "GIA-X19_co_mat_khau.xlsx")
    res["GH-07d_xlsm"] = upload((FIX / "GIA-X19_khong_co_vba.xlsm").read_bytes(), "GIA-X19_khong_co_vba.xlsm")
    long_name = ("Biểu mẫu nội kiểm khoa xét nghiệm tháng chín năm hai nghìn không trăm hai mươi sáu # & ' bản giả " * 3)[:200].strip() + ".xlsx"
    named = upload(qc, long_name)
    if named["document_id"]:
        _, headers, raw, _ = call("GET", f"/api/documents/{named['document_id']}/export")
        disposition = headers.get("Content-Disposition") or headers.get("content-disposition") or ""
        named["export_name"] = urllib.parse.unquote(disposition.split("UTF-8''")[-1])
        named["export_bytes"] = len(raw)
    named["name_sent"] = long_name
    res["GH-07e_ten_dai"] = named

    # GH-02 qua API: hai lần lưu cùng dựa trên bản 1.
    doc = upload(qc, "GIA-QC-01_hai_tab.xlsx")["document_id"]
    first = post_json(f"/api/documents/{doc}/versions", {"base_version": 1, "author": "Người Giả", "edits": [{"sheet": SHEET, "cell": "B8", "value": 5.6}]})
    second = post_json(f"/api/documents/{doc}/versions", {"base_version": 1, "author": "Người Giả", "edits": [{"sheet": SHEET, "cell": "B9", "value": 5.7}]})
    res["GH-02_api"] = {"first": first[0], "second": second[0], "second_body": body_of(second[2])}

    # GH-06: hơn 5.000 ô trong một lần lưu.
    doc6 = upload(qc, "GIA-QC-01_nhieu_o.xlsx")["document_id"]
    many = [{"sheet": SHEET, "cell": f"{col}{row}", "value": row} for row in range(100, 200) for col in ("AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY")]
    s5000 = post_json(f"/api/documents/{doc6}/versions", {"base_version": 1, "author": "Người Giả", "edits": many[:5000]})
    s5001 = post_json(f"/api/documents/{doc6}/versions", {"base_version": 2 if s5000[0] == 201 else 1, "author": "Người Giả", "edits": many[:5001]})
    res["GH-06"] = {"edits_5000": {"status": s5000[0], "seconds": s5000[3]}, "edits_5001": {"status": s5001[0], "seconds": s5001[3], "body": body_of(s5001[2])}}

    # GH-09: lịch sử dài. Một ô đổi mỗi lần lưu.
    before = store_bytes()
    doc9 = upload(qc, "GIA-QC-01_lich_su_dai.xlsx")["document_id"]
    _, _, grid_raw, _ = call("GET", f"/api/documents/{doc9}/grid")
    saves = []
    for n in range(1, 101):
        status, _, _, seconds = post_json(f"/api/documents/{doc9}/versions", {"base_version": n, "author": "Người Giả", "edits": [{"sheet": SHEET, "cell": "B8", "value": 5 + n / 1000}]})
        saves.append({"version": n + 1, "status": status, "seconds": seconds})
        if n in (50, 100):
            _, _, raw, list_seconds = call("GET", f"/api/documents/{doc9}/versions")
            res[f"GH-09_list_after_{n}"] = {"seconds": list_seconds, "count": len(body_of(raw)) if raw.startswith(b"[") else body_of(raw)}
    after = store_bytes()
    status, _, raw, seconds = post_json(f"/api/documents/{doc9}/versions/1/restore", {"base_version": 101, "author": "Người Giả"})
    res["GH-09"] = {
        "save_status": sorted({s["status"] for s in saves}),
        "save_seconds_first10_median": sorted(s["seconds"] for s in saves[:10])[5],
        "save_seconds_last10_median": sorted(s["seconds"] for s in saves[-10:])[5],
        "save_seconds_max": max(s["seconds"] for s in saves),
        "store_bytes_before": before,
        "store_bytes_after": after,
        "file_bytes": len(qc),
        "restore_status": status,
        "restore_seconds": seconds,
        "restore_version": body_of(raw).get("version"),
        "grid_bytes": len(grid_raw),
    }
    res["finished"] = time.strftime("%Y-%m-%d %H:%M:%S")
    OUT.mkdir(parents=True, exist_ok=True)
    (OUT / "limits.json").write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf8")
    print(json.dumps(res, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
