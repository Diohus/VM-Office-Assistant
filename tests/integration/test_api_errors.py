"""Q124: lỗi kiểm dữ liệu của API trả câu tiếng Việt và không bao giờ trả lại giá trị đã gửi (số trong ô)."""
from __future__ import annotations

import json

from conftest import FIX, request, upload

SHEET = "KQ_QC_Thang"
# Giá trị dấu vết: nếu xuất hiện trong phản hồi lỗi thì API đã trả lại dữ liệu người dùng gửi.
SECRET = 987654.321123
SECRET_TEXT = "DAU-VET-KHONG-DUOC-TRA-LAI"


def post(path: str, payload: dict):
    status, _, raw = request("POST", path, json.dumps(payload).encode(), headers={"Content-Type": "application/json"})
    return status, raw


def doc() -> str:
    status, body = upload((FIX / "GIA-QC-01_mau.xlsx").read_bytes(), "GIA-QC-01_mau.xlsx")
    assert status == 201, body
    return body["document_id"]


def assert_no_echo(raw: bytes) -> dict:
    text = raw.decode("utf8")
    assert "987654" not in text and SECRET_TEXT not in text, "phản hồi lỗi không được chứa giá trị đã gửi"
    assert '"input"' not in text and '"ctx"' not in text
    return json.loads(text)


def test_qua_5000_o_bao_tieng_viet_khong_tra_lai_gia_tri():
    edits = [{"sheet": SHEET, "cell": f"AA{r}", "value": SECRET} for r in range(1, 5002)]
    status, raw = post(f"/api/documents/{doc()}/versions", {"base_version": 1, "author": "Người Giả", "edits": edits})
    body = assert_no_echo(raw)
    assert status == 413 and body["code"] == "EDITS_TOO_MANY"
    assert "quá 5000 ô sửa" in body["message"] and "chưa ghi" in body["message"]
    assert len(raw) < 2000, "không gửi lại cả danh sách ô"


def test_dung_5000_o_van_luu_duoc():
    edits = [{"sheet": SHEET, "cell": f"AA{r}", "value": r} for r in range(1, 5001)]
    status, raw = post(f"/api/documents/{doc()}/versions", {"base_version": 1, "author": "Người Giả", "edits": edits})
    assert status == 201, raw[:300]


def test_loi_kiem_du_lieu_khac_khong_tra_lai_gia_tri():
    # Giá trị ô sai kiểu (một đối tượng) và tên ô quá dài.
    bad_value = {"sheet": SHEET, "cell": "B8", "value": {"so": SECRET, "chu": SECRET_TEXT}}
    bad_cell = {"sheet": SHEET, "cell": SECRET_TEXT, "value": SECRET}
    status, raw = post(f"/api/documents/{doc()}/versions", {"base_version": 1, "author": "Người Giả", "edits": [bad_value, bad_cell]})
    body = assert_no_echo(raw)
    assert status == 422 and body["code"] == "INVALID_REQUEST" and "không bị đổi" in body["message"]
    assert {f["field"] for f in body["detail"]["fields"]} >= {"body.edits.1.cell"}


def test_pdf_dau_thap_phan_la_bi_tu_choi_bang_cau_tieng_viet():
    status, raw = post(f"/api/documents/{doc()}/pdf", {"grid_values": [{"sheet": SHEET, "cell": "B39", "value": SECRET}], "decimal": SECRET_TEXT})
    body = assert_no_echo(raw)
    assert status == 422 and body["code"] == "INVALID_REQUEST"
