"""T-DEV-09 mốc 1: API lát 1 nối AI (Q133). Đối chiếu docs/product/acceptance_lat1.md (mã dòng trong tên test).

Chạy trên bộ thử 8089. Đọc thẳng cơ sở dữ liệu của bộ thử qua docker exec để kiểm không có tên người, giá trị người bệnh.
"""
from __future__ import annotations

import json
import pathlib
import subprocess
import sys
import threading
import urllib.parse

import pytest

from conftest import FIX, PROJECT, request, upload

ROOT = pathlib.Path(__file__).resolve().parents[2]
SRC = FIX / "nguon_lat1"
SHEET = "KQ_QC_Thang"
# Chuỗi không được nằm trong lô lưu, bảng nguồn, hay phản hồi API nào (2.4, 2.11).
PEOPLE = ["Bệnh Nhân Giả", "BN-GIA-", "Trần Giả Hai", "Nguyễn Giả Một", "Lê Giả A"]


def post(path: str, payload: dict | None = None, data: bytes | None = None, method: str = "POST"):
    body = data if data is not None else (json.dumps(payload).encode() if payload is not None else None)
    headers = {"Content-Type": "application/json"} if payload is not None else {}
    status, _, raw = request(method, path, body, headers=headers)
    return status, raw, (json.loads(raw) if raw and raw[:1] in (b"{", b"[") else None)


def qc_doc(name: str = "GIA-QC-01_mau.xlsx") -> dict:
    status, doc = upload((FIX / name).read_bytes(), name)
    assert status == 201, doc
    return doc


def add_source(doc: dict, path: pathlib.Path):
    return post(f"/api/documents/{doc['document_id']}/sources?filename={urllib.parse.quote(path.name)}", data=path.read_bytes())


def new_batch(doc: dict, path: pathlib.Path, base_version: int = 1) -> tuple[int, bytes, dict]:
    status, raw, src = add_source(doc, path)
    assert status == 201, raw[:300]
    return post(f"/api/documents/{doc['document_id']}/fill", {"source_id": src["source_id"], "base_version": base_version})


def opened(doc: dict, path: pathlib.Path) -> dict:
    status, raw, b = new_batch(doc, path)
    assert status == 201, raw[:400]
    status, raw, b = post(f"/api/fill/{b['batch_id']}/open")
    assert status == 200, raw[:400]
    return b


def apply(batch: dict, decisions: list[dict], base_version: int, document_id: str | None = None):
    body = {"document_id": batch["document_id"] if document_id is None else document_id, "base_version": base_version, "author": "Người Giả", "decisions": decisions}
    return post(f"/api/fill/{batch['batch_id']}/apply", body)


def changes(batch: dict, color: str | None = None) -> list[dict]:
    return [c for c in batch["batch"]["changes"] if color is None or c["confidence"]["final_color"] == color]


def grid_value(doc_id: str, a1: str, version: int | None = None):
    q = f"?version={version}" if version else ""
    _, _, raw = request("GET", f"/api/documents/{doc_id}/grid{q}")
    sheet = next(s for s in json.loads(raw)["sheets"] if s["name"] == SHEET)
    col = ord(a1[0]) - 65
    row = int(a1[1:]) - 1
    cell = next((c for c in sheet["cells"] if c["r"] == row and c["c"] == col), {})
    return cell.get("v")


def sql(query: str) -> str:
    out = subprocess.run(["docker", "exec", f"{PROJECT}-db-1", "psql", "-U", "vm", "-d", "vm", "-At", "-c", query],
                         capture_output=True, text=True, encoding="utf8", timeout=60)
    assert out.returncode == 0, out.stderr
    return out.stdout


# ---- 1. Bắt đầu ----
def test_chi_bieu_mau_qc_mau_moi_dien_duoc_bao_truoc_khi_doc():
    assert qc_doc()["fill_supported"] is True
    other = qc_doc("GIA-QC-01_da_dien.xlsx")
    assert other["fill_supported"] is False
    status, raw, body = add_source(other, FIX / "GIA-NGUON_2026-09.xlsx")
    assert status == 422 and body["code"] == "FILL_TEMPLATE_UNSUPPORTED" and "chỉ điền được biểu mẫu QC mẫu" in body["message"]


def test_1_6_nguon_hong_qua_cong_dau_vao():
    doc = qc_doc()
    status, _, body = add_source(doc, FIX / "GIA-X19_hong.xlsx")
    assert status == 422 and body["code"] == "FILE_CORRUPT" and "không bị đổi" in body["message"]
    status, _, body = add_source(doc, FIX / "GIA-QC-01_da_dien_doi_cu.xls")
    assert status == 422 and body["code"] == "FILL_SOURCE_XLSX_ONLY"


# ---- 2. Đề xuất và màu ----
def test_2_1_2_2_2_3_x01_dung_o_khong_xanh():
    doc = qc_doc()
    status, raw, b = new_batch(doc, FIX / "GIA-NGUON_2026-09.xlsx")
    assert status == 201, raw[:400]
    assert b["status"] == "draft" and b["batch"]["schema_version"] == "0.4.1"
    cells = [c["target"]["a1_at_write"] for c in changes(b)]
    assert len(cells) == 180
    assert not changes(b, "green")
    assert all(a1[0] in "BCDEFG" and 8 <= int(a1[1:]) <= 37 for a1 in cells), "chỉ B8:G37, không cột A, H, I, không dòng công thức"


def test_2_4_2_11_khong_ten_nguoi_khong_gia_tri_nguoi_benh():
    doc = qc_doc()
    texts = []
    for name in ("X17_cot_nguoi_benh_gia.xlsx", "nguoi_chay_ten_gia.xlsx"):
        status, raw, b = new_batch(doc, SRC / name)
        assert status == 201, raw[:300]
        texts.append(raw.decode("utf8"))
        status, raw, _ = post(f"/api/fill/{b['batch_id']}/open")
        texts.append(raw.decode("utf8"))
        _, raw, _ = post(f"/api/fill/{b['batch_id']}", method="GET")
        texts.append(raw.decode("utf8"))
        reasons = {n["reason"] for n in b["batch"]["notices"]}
        assert "PERSON_COLUMN_LEFT_BLANK" in reasons
        if name.startswith("X17"):
            assert "PHI_COLUMN_SKIPPED" in reasons
        assert all(c["target"]["a1_at_write"][0] != "H" for c in changes(b)), "cột Người chạy không có đề xuất"
    _, raw, _ = post(f"/api/documents/{doc['document_id']}/fill", method="GET")
    texts.append(raw.decode("utf8"))
    texts.append(sql(f"SELECT batch::text, targets::text, decisions::text FROM fill_batch WHERE document_id = '{doc['document_id']}'"))
    texts.append(sql(f"SELECT name, sha256 FROM source_file WHERE document_id = '{doc['document_id']}'"))
    for text in texts:
        for secret in PEOPLE:
            assert secret not in text, f"{secret} không được nằm trong lô, bảng nguồn hay phản hồi API"


# ---- 3, 4. Duyệt và ghi ----
def test_4_2_4_3_4_6_ghi_chi_lay_so_tu_lo():
    doc = qc_doc()
    b = opened(doc, FIX / "GIA-NGUON_2026-09.xlsx")
    yellow = changes(b, "yellow")
    take = yellow[:30]
    # Gửi kèm số tự đặt: phải bị bỏ qua, số ghi là số trong lô (4.6).
    status, raw, r = apply(b, [{"change_id": c["change_id"], "action": "accept", "value": 999} for c in take], 1)
    assert status == 200, raw[:300]
    assert r["written"] == 30 and r["version"]["version"] == 2 and r["review_again"] == {}
    assert r["version"]["summary"] == "Nhận 30 ô từ file GIA-NGUON_2026-09.xlsx", "không có ô đỏ thì không có phần sau dấu chấm phẩy"
    first = take[0]
    assert grid_value(doc["document_id"], first["target"]["a1_at_write"]) == float(first["value"]["normalized"])
    _, _, versions = post(f"/api/documents/{doc['document_id']}/versions", method="GET")
    assert versions[0]["author"] == "Người Giả" and versions[-1]["summary"] == "Tải lên"


def test_4_6_tu_choi_id_la_o_do_va_ban_cu():
    doc = qc_doc()
    b = opened(doc, SRC / "X10_so_vo_ly.xlsx")
    red = changes(b, "red")[0]
    status, _, body = apply(b, [{"change_id": red["change_id"], "action": "accept"}], 1)
    assert status == 422 and body["code"] == "FILL_DECISION_INVALID" and "Ô đỏ" in body["message"]
    status, _, body = apply(b, [{"change_id": "c-khong-co", "action": "accept"}], 1)
    assert status == 422 and body["code"] == "FILL_DECISION_INVALID"
    status, _, body = apply(b, [{"change_id": changes(b, "yellow")[0]["change_id"], "action": "accept"}], 7)
    assert status == 409 and body["code"] == "STALE_VERSION"
    assert len(post(f"/api/documents/{doc['document_id']}/versions", method="GET")[2]) == 1, "không ghi gì"


def test_ghi_lo_cua_tai_lieu_khac_bi_tu_choi():
    # PO/BA 11/11: mở tài liệu B (cùng tên, cùng bản 1) khi khung của A còn mở, Ghi đã ghi vào A. Máy chủ phải từ chối.
    a = qc_doc()
    b_doc = qc_doc()
    b = opened(a, SRC / "X10_so_vo_ly.xlsx")
    take = [{"change_id": changes(b, "yellow")[0]["change_id"], "action": "accept"}]
    status, _, body = apply(b, take, 1, document_id=b_doc["document_id"])
    assert status == 409 and body["code"] == "FILL_WRONG_DOCUMENT" and "biểu mẫu khác" in body["message"]
    status, _, body = apply(b, take, 1, document_id="")
    assert status == 422
    for d in (a, b_doc):
        assert len(post(f"/api/documents/{d['document_id']}/versions", method="GET")[2]) == 1, "không tài liệu nào bị ghi"
    status, _, r = apply(b, take, 1)
    assert status == 200 and r["written"] == 1, "đúng tài liệu thì vẫn ghi được"


def test_d4_o_do_chon_lan_doc_tu_go_de_trong_3_8_3_8a():
    doc = qc_doc()
    b = opened(doc, SRC / "X04_mot_ngay_hai_lan_chay.xlsx")
    red = changes(b, "red")
    assert len(red) == 1 and red[0]["value"] is None and len(red[0]["evidence"]) == 2
    dup = red[0]
    yellow = changes(b, "yellow")
    # Ghi ô vàng, ô đỏ còn chờ (3.8): lô vẫn mở, vẫn tính D10.
    status, _, r = apply(b, [{"change_id": c["change_id"], "action": "accept"} for c in yellow], 1)
    assert status == 200 and r["pending"] == 1
    _, _, again = post(f"/api/fill/{b['batch_id']}", method="GET")
    assert again["status"] == "open" and again["decisions"][dup["change_id"]]["state"] == "pending"
    # Tự gõ số nhập nhằng: từ chối như khi gõ trong lưới.
    status, _, body = apply(b, [{"change_id": dup["change_id"], "action": "red_typed", "text": "1,234"}], 2)
    assert status == 422 and "hiểu hai cách" in body["message"]
    status, _, body = apply(b, [{"change_id": dup["change_id"], "action": "red_typed", "text": "1,5,5"}], 2)
    assert status == 422 and "không phải một số rõ ràng" in body["message"]
    # Chọn lần đọc 2: số lấy từ evidence[1].normalized.
    status, _, r = apply(b, [{"change_id": dup["change_id"], "action": "red_choice", "index": 1}], 2)
    assert status == 200 and r["pending"] == 0 and r["version"]["version"] == 3
    assert r["version"]["summary"].endswith("; ô đỏ: 1 chọn lần đọc 2")
    assert grid_value(doc["document_id"], dup["target"]["a1_at_write"]) == float(dup["evidence"][1]["normalized"])
    assert post(f"/api/fill/{b['batch_id']}", method="GET")[2]["status"] == "done"


def test_d4_tu_go_va_de_trong():
    doc = qc_doc()
    b = opened(doc, SRC / "X10_so_vo_ly.xlsx")
    red = changes(b, "red")
    assert len(red) == 5
    decisions = [{"change_id": red[0]["change_id"], "action": "red_typed", "text": "5.5"},
                 {"change_id": red[1]["change_id"], "action": "red_blank"},
                 {"change_id": red[2]["change_id"], "action": "red_blank"}]
    status, _, r = apply(b, decisions, 1)
    assert status == 200 and r["written"] == 1 and r["pending"] == 2 + len(changes(b, "yellow"))
    assert r["version"]["summary"] == "Nhận 1 ô từ file X10_so_vo_ly.xlsx; ô đỏ: 1 tự gõ, 2 để trống"
    assert grid_value(doc["document_id"], red[0]["target"]["a1_at_write"]) == 5.5
    assert grid_value(doc["document_id"], red[1]["target"]["a1_at_write"]) is None


def test_3_8a_o_bi_doi_tu_cho_khac_phai_duyet_lai():
    doc = qc_doc()
    b = opened(doc, FIX / "GIA-NGUON_2026-09.xlsx")
    c = changes(b, "yellow")[0]
    a1 = c["target"]["a1_at_write"]
    status, _, _ = post(f"/api/documents/{doc['document_id']}/versions", {"base_version": 1, "author": "Người Giả", "edits": [{"sheet": SHEET, "cell": a1, "value": 7.77}]})
    assert status == 201
    status, _, r = apply(b, [{"change_id": c["change_id"], "action": "accept"}], 2)
    assert status == 200 and r["written"] == 0 and r["review_again"] == {c["change_id"]: 7.77}, "không ghi đè im lặng"
    assert grid_value(doc["document_id"], a1) == 7.77
    status, _, r = apply(b, [{"change_id": c["change_id"], "action": "accept", "seen": "7.77"}], 2)
    assert status == 200 and r["written"] == 1 and grid_value(doc["document_id"], a1) == float(c["value"]["normalized"])


# ---- 5. Một lô mỗi ô (D10) ----
def test_5_3_lo_thu_hai_cham_o_dang_cho_bi_tu_choi():
    doc = qc_doc()
    opened(doc, FIX / "GIA-NGUON_2026-09.xlsx")
    status, raw, b2 = new_batch(doc, SRC / "X07_so_dang_chu_dau_phay.xlsx")
    assert status == 201, raw[:300]
    status, _, body = post(f"/api/fill/{b2['batch_id']}/open")
    assert status == 409 and body["code"] == "FILL_CONFLICT" and "đang chờ duyệt" in body["message"] and body["detail"]["cells"]


def test_5_3_hai_tab_cung_luc_chi_mot_lo_mo():
    doc = qc_doc()
    _, _, src1 = add_source(doc, FIX / "GIA-NGUON_2026-09.xlsx")
    _, _, src2 = add_source(doc, SRC / "X07_so_dang_chu_dau_phay.xlsx")
    # Hai lô cùng tạo rồi cùng mở song song: tạo lô mới xóa lô nháp cũ (6.1), nên tab còn lại phải thấy lô của mình mất hoặc bị D10 chặn.
    _, _, b1 = post(f"/api/documents/{doc['document_id']}/fill", {"source_id": src1["source_id"], "base_version": 1})
    _, _, b2 = post(f"/api/documents/{doc['document_id']}/fill", {"source_id": src2["source_id"], "base_version": 1})
    results = {}

    def open_it(key, b):
        results[key] = post(f"/api/fill/{b['batch_id']}/open")[0]

    threads = [threading.Thread(target=open_it, args=(k, b)) for k, b in (("a", b1), ("b", b2))]
    for t in threads:
        t.start()
    for t in threads:
        t.join()
    assert sorted(results.values()).count(200) == 1, results
    assert int(sql(f"SELECT count(*) FROM fill_batch WHERE document_id = '{doc['document_id']}' AND status = 'open'")) == 1


# ---- N5: bỏ lần điền đang mở; N8: lô 0 ô ----
def test_n5_bo_lan_dien_dang_mo_giu_o_da_ghi_va_cho_lo_moi():
    doc = qc_doc()
    b = opened(doc, FIX / "GIA-NGUON_2026-09.xlsx")
    first = changes(b, "yellow")[0]
    status, _, r = apply(b, [{"change_id": first["change_id"], "action": "accept"}], 1)
    assert status == 200 and r["written"] == 1
    pending = r["pending"]
    status, _, r = post(f"/api/fill/{b['batch_id']}", method="DELETE")
    assert status == 200 and r == {"deleted": False, "dropped": pending}
    assert post(f"/api/documents/{doc['document_id']}/fill", method="GET")[2] == [], "không còn câu Còn một lần điền"
    assert len(post(f"/api/documents/{doc['document_id']}/versions", method="GET")[2]) == 2, "bỏ không tạo bản mới"
    assert grid_value(doc["document_id"], first["target"]["a1_at_write"]) is not None, "ô đã ghi giữ nguyên"
    status, _, body = apply(b, [{"change_id": changes(b, "yellow")[1]["change_id"], "action": "accept"}], 2)
    assert status == 409 and body["code"] == "FILL_GONE"
    # Lô mới chạm cùng các ô mở được ngay (D10 không còn giữ ô).
    status, raw, b2 = new_batch(doc, FIX / "GIA-NGUON_2026-09.xlsx", base_version=2)
    assert status == 201, raw[:300]
    assert post(f"/api/fill/{b2['batch_id']}/open")[0] == 200
    # Gọi lại lần nữa: không lỗi, không đổi gì.
    assert post(f"/api/fill/{b['batch_id']}", method="DELETE")[2] == {"deleted": False, "dropped": 0}


def test_n8_lo_0_o_ket_thuc_ngay_khong_bao_con_o_cho():
    doc = qc_doc()
    status, raw, b = new_batch(doc, SRC / "X13_sai_thang_10.xlsx")
    assert status == 201 and b["batch"]["changes"] == [], raw[:300]
    status, _, b = post(f"/api/fill/{b['batch_id']}/open")
    assert status == 200 and b["status"] == "done"
    assert post(f"/api/documents/{doc['document_id']}/fill", method="GET")[2] == []


# ---- 6. Hủy ----
def test_6_1_huy_xoa_lo_nhap_va_khong_chan_lo_moi():
    doc = qc_doc()
    status, _, b = new_batch(doc, FIX / "GIA-NGUON_2026-09.xlsx")
    assert status == 201
    status, _, r = post(f"/api/fill/{b['batch_id']}", method="DELETE")
    assert status == 200 and r["deleted"] is True
    assert int(sql(f"SELECT count(*) FROM fill_batch WHERE document_id = '{doc['document_id']}'")) == 0
    # Lô nháp bỏ dở (yêu cầu hủy không tới): tạo lô mới thì lô cũ bị xóa, lô mới mở được.
    _, _, stale = new_batch(doc, FIX / "GIA-NGUON_2026-09.xlsx")
    _, _, fresh = new_batch(doc, FIX / "GIA-NGUON_2026-09.xlsx")
    assert post(f"/api/fill/{stale['batch_id']}", method="GET")[0] == 404
    assert post(f"/api/fill/{fresh['batch_id']}/open")[0] == 200


# ---- Lớp hai của backend: kiểm lô, không cần container ----
def _fill_module():
    sys.path.insert(0, str(ROOT / "backend"))
    try:
        for name in [n for n in sys.modules if n == "app" or n.startswith("app.")]:
            del sys.modules[name]
        import app.fill as fill
        return fill
    finally:
        sys.path.remove(str(ROOT / "backend"))


def _grid():
    cells = [{"r": 37, "c": 1, "f": "=SUM(B8:B37)"}]
    return {"sheets": [{"name": SHEET, "rows": 60, "cols": 16, "cells": cells, "merges": [[5, 1, 5, 2]], "protected": False}, {"name": "An", "rows": 5, "cols": 5, "cells": [], "merges": []}]}


def _batch(**change):
    base = {"change_id": "c-1", "op": "set_value", "target": {"sheet": SHEET, "a1_at_write": "B8"}, "previous": {"value": None, "formula": None},
            "value": {"type": "number", "normalized": "5.51", "raw_text": "5.51"}, "evidence": [], "confidence": {"final_color": "yellow"}}
    base.update(change)
    return {"schema_version": "0.4", "template": {"spec_id": "GIA-QC-01"}, "changes": [base]}


@pytest.mark.parametrize("change,why", [
    ({"confidence": {"final_color": "green"}}, "lát 1 không có ô xanh"),
    ({"target": {"sheet": SHEET, "a1_at_write": "B38"}}, "ô công thức"),
    ({"target": {"sheet": SHEET, "a1_at_write": "C6"}}, "trong ô gộp, không phải ô đầu"),
    ({"target": {"sheet": "An", "a1_at_write": "A1"}}, "sheet ẩn"),
    ({"target": {"sheet": SHEET, "a1_at_write": "Z999"}}, "ngoài bảng"),
    ({"value": None}, "ô vàng không có số"),
])
def test_lo_sai_bi_tu_choi_ca_lo(change, why):
    fill = _fill_module()
    with pytest.raises(fill.BatchInvalid):
        fill.check_batch(_batch(**change), _grid(), ["An"])


def test_lo_dung_qua_va_lo_cu_bi_tu_choi():
    fill = _fill_module()
    targets, decisions = fill.check_batch(_batch(), _grid(), ["An"])
    assert targets == {"c-1": f"{SHEET}!B8"} and decisions == {"c-1": {"state": "pending"}}
    old = _batch()
    old["schema_version"] = "0.3"
    with pytest.raises(fill.BatchInvalid):
        fill.check_batch(old, _grid(), ["An"])


def _with_items(items, version="0.4.1"):
    b = _batch()
    b["schema_version"] = version
    b["notices"] = [{"reason": "SOURCE_MISSING_DAYS", "count": len(items), "detail": "x", "items": items}]
    return b


def test_items_041_qua_va_lo_04_khong_co_items_van_doc_duoc():
    fill = _fill_module()
    fill.check_batch(_with_items(["2026-09-08", "2026-09-18"]), _grid(), ["An"])
    fill.check_batch(_with_items([] , "0.4"), _grid(), ["An"])
    old = _batch()
    old["notices"] = [{"reason": "SOURCE_MISSING_DAYS", "count": 1, "detail": "Nguồn thiếu: 2026-09-08."}]
    fill.check_batch(old, _grid(), ["An"])
    fill.check_batch(_with_items(["x" * 120] * 200), _grid(), ["An"])


@pytest.mark.parametrize("items,why", [
    (["2026-09-08"] * 201, "quá 200 phần tử"),
    (["x" * 121], "phần tử quá 120 ký tự"),
    ([5], "phần tử không phải chữ"),
    ("2026-09-08", "không phải danh sách"),
])
def test_items_vuot_gioi_han_bi_tu_choi_ca_lo(items, why):
    fill = _fill_module()
    with pytest.raises(fill.BatchInvalid):
        fill.check_batch(_with_items(items), _grid(), ["An"])


def test_hieu_so_go_tay_giong_ban_trinh_duyet():
    """Bản Python của typedNumber.js: cùng bộ ca với typed_number.test.mjs."""
    p = _fill_module().parse_typed_number
    for text, value in [("5,51", 5.51), ("5.51", 5.51), ("-5,5", -5.5), ("0", 0), ("12", 12), ("0,125", 0.125), ("1234,567", 1234.567),
                        ("1.234,5", 1234.5), ("1,234.5", 1234.5), ("1.234.567", 1234567), ("1,5", 1.5), ("5,5", 5.5), ("12,5", 12.5), ("-2,75", -2.75)]:
        assert p(text) == {"kind": "number", "value": value}, text
    for text in ["1.234", "1,234", "12,345", "999.000", "-1.234"]:
        assert p(text)["kind"] == "ambiguous", text
    for text in ["1,2,3", "5,5.1", "5,5511.234", "1.2.3", "1,,2", ",5", "5,", "1.234.56", "1,5,5"]:
        assert p(text)["kind"] == "unclear", text
    for text in ["abc", "", "=1+1", "01/09", "5,5a"]:
        assert p(text)["kind"] == "other", text


# ---- Đua thật: hai lần áp cùng lô, cùng base_version (PM, trước commit mốc 1) ----
def _race(b: dict, first: list[dict], second: list[dict]) -> list[tuple[int, dict]]:
    out = [None, None]
    gate = threading.Barrier(2)

    def go(i, decisions):
        gate.wait()
        status, _, body = apply(b, decisions, 1)
        out[i] = (status, body)

    threads = [threading.Thread(target=go, args=(0, first)), threading.Thread(target=go, args=(1, second))]
    for t in threads:
        t.start()
    for t in threads:
        t.join()
    return out


@pytest.mark.parametrize("second_writes", [True, False], ids=["ca_hai_cung_ghi", "mot_ghi_mot_bo"])
def test_hai_lan_ap_cung_luc_chi_mot_lan_duoc_va_khong_mat_quyet_dinh(second_writes):
    doc = qc_doc()
    b = opened(doc, FIX / "GIA-NGUON_2026-09.xlsx")
    yellow = changes(b, "yellow")
    first = [{"change_id": c["change_id"], "action": "accept"} for c in yellow[:10]]
    second = [{"change_id": c["change_id"], "action": "accept" if second_writes else "reject"} for c in yellow[10:20]]
    results = _race(b, first, second)
    ok = [r for r in results if r[0] == 200]
    refused = [r for r in results if r[0] != 200]
    assert len(ok) == 1 and len(refused) == 1, [(s, (body or {}).get("code")) for s, body in results]
    assert refused[0][0] == 409 and refused[0][1]["code"] in ("STALE_VERSION", "FILL_BATCH_CHANGED"), refused[0][1]
    versions = post(f"/api/documents/{doc['document_id']}/versions", method="GET")[2]
    winner = ok[0][1]
    assert len(versions) == (2 if winner["written"] else 1), "đúng một bản mới khi lần thắng có ghi"
    # Không mất quyết định: trạng thái trong cơ sở dữ liệu đúng bằng trạng thái lần thắng trả về.
    stored = json.loads(sql(f"SELECT decisions::text FROM fill_batch WHERE batch_id = '{b['batch_id']}'"))
    assert stored == winner["decisions"]
    loser_ids = {d["change_id"] for d in (second if results[0][0] == 200 else first)}
    assert all(stored[cid]["state"] == "pending" for cid in loser_ids), "lần bị từ chối không để lại dấu gì"


_DB_RACE = r'''
import json, sys, threading, uuid
sys.path.insert(0, "/opt/app")
from app import db
doc, batch, cid_a, cid_b = sys.argv[1:5]
row = db.batch(uuid.UUID(batch))
expected = row["decisions"]
gate = threading.Barrier(2)
out = {}
def go(name, cid):
    decisions = {k: dict(v) for k, v in expected.items()}
    decisions[cid] = {"state": "written", "how": "proposal", "version": None}
    write = {"sha": "0" * 64, "size": 1, "meta": {}, "author": "Người Giả", "summary": "đua", "edits": [{"sheet": "KQ_QC_Thang", "cell": "Z1", "value": 1}]}
    gate.wait()
    try:
        r = db.apply_batch(uuid.UUID(batch), uuid.UUID(doc), 1, decisions, write, expected)
        out[name] = "ok:" + str(r["version"])
    except Exception as exc:
        out[name] = type(exc).__name__
ts = [threading.Thread(target=go, args=("a", cid_a)), threading.Thread(target=go, args=("b", cid_b))]
[t.start() for t in ts]
[t.join() for t in ts]
print(json.dumps(out))
'''


def test_dua_that_trong_co_so_du_lieu_mot_ban_moi_lan_kia_la_ban_lech():
    """Hai giao dịch apply_batch đua thật (không qua reader): đúng một bản mới, lần kia StaleVersion hoặc BatchChanged, không bao giờ lỗi khác (500)."""
    doc = qc_doc()
    b = opened(doc, FIX / "GIA-NGUON_2026-09.xlsx")
    yellow = changes(b, "yellow")
    out = subprocess.run(["docker", "exec", "-i", f"{PROJECT}-api-1", "python", "-", doc["document_id"], b["batch_id"], yellow[0]["change_id"], yellow[1]["change_id"]],
                         input=_DB_RACE, capture_output=True, text=True, encoding="utf8", timeout=60)
    assert out.returncode == 0, out.stderr[-1500:]
    result = json.loads(out.stdout.strip().splitlines()[-1])
    assert sorted(v.split(":")[0] for v in result.values()) in (["BatchChanged", "ok"], ["StaleVersion", "ok"]), result
    assert int(sql(f"SELECT count(*) FROM document_version WHERE document_id = '{doc['document_id']}'")) == 2
