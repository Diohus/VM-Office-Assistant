"""Ghi ô, phiên bản, khôi phục qua API thật. Bản xuất được đối chiếu với Excel thật ở tools/excel_check/q48_check.py."""
from __future__ import annotations

import io
import json
import zipfile

import openpyxl

from conftest import FIX, PROJECT, request, upload

SHEET = "KQ_QC_Thang"


def post(path: str, body: dict):
    status, _, raw = request("POST", path, json.dumps(body).encode("utf8"), headers={"Content-Type": "application/json"})
    return status, json.loads(raw)


def new_doc(name: str = "GIA-QC-01_da_dien.xlsx") -> dict:
    status, doc = upload((FIX / name).read_bytes(), name)
    assert status == 201
    return doc


def save(doc: dict, edits: list[dict], base: int | None = None):
    return post(f"/api/documents/{doc['document_id']}/versions", {"base_version": base or doc["version"], "author": "Người Giả", "edits": edits})


def export(doc: dict, version: int | None = None) -> bytes:
    suffix = f"?version={version}" if version else ""
    return request("GET", f"/api/documents/{doc['document_id']}/export{suffix}")[2]


def test_ghi_tao_ban_moi_va_ban_goc_khong_doi():
    doc = new_doc()
    original = export(doc)
    status, v2 = save(doc, [{"sheet": SHEET, "cell": "B8", "value": 5.5}])
    assert status == 201 and v2["version"] == 2 and v2["summary"] == "Sửa 1 ô" and v2["author"] == "Người Giả"
    assert export(doc, 1) == original
    ws = openpyxl.load_workbook(io.BytesIO(export(doc)))[SHEET]
    assert ws["B8"].value == 5.5
    grid = json.loads(request("GET", f"/api/documents/{doc['document_id']}/grid")[2])
    cells = {(c["r"], c["c"]): c for c in grid["sheets"][0]["cells"]}
    assert cells[(7, 1)]["v"] == 5.5


def test_ban_xuat_khong_con_so_cu_va_co_ma_dinh_danh():
    doc = new_doc()
    save(doc, [{"sheet": SHEET, "cell": "B8", "value": 500}])
    data = export(doc)
    values = openpyxl.load_workbook(io.BytesIO(data), data_only=True)[SHEET]
    assert values["B39"].value is None
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        assert doc["document_id"].encode() in z.read("docProps/custom.xml")
        assert b'fullCalcOnLoad="1"' in z.read("xl/workbook.xml")


def test_ban_cu_khong_ghi_de_ban_moi():
    doc = new_doc()
    assert save(doc, [{"sheet": SHEET, "cell": "B8", "value": 1}])[0] == 201
    status, body = save(doc, [{"sheet": SHEET, "cell": "B9", "value": 2}], base=1)
    assert status == 409 and body["code"] == "STALE_VERSION" and "không bị đổi" in body["message"]


def test_khoi_phuc_tao_ban_moi_lich_su_con_nguyen():
    doc = new_doc()
    original = export(doc)
    save(doc, [{"sheet": SHEET, "cell": "B8", "value": 1}])
    status, v3 = post(f"/api/documents/{doc['document_id']}/versions/1/restore", {"base_version": 2, "author": "Người Giả"})
    assert status == 201 and v3["version"] == 3 and v3["summary"] == "Khôi phục bản 1"
    assert export(doc) == original
    history = json.loads(request("GET", f"/api/documents/{doc['document_id']}/versions")[2])
    assert [h["version"] for h in history] == [3, 2, 1] and history[0]["restored_from"] == 1


def test_o_khoa_bi_tu_choi_khong_tao_ban_moi():
    doc = new_doc("GIA-X19_sheet_khoa.xlsx")
    status, body = save(doc, [{"sheet": SHEET, "cell": "B8", "value": 1}])
    assert status == 422 and body["code"] == "WRITE_REFUSED" and body["detail"]["cells"][0]["cell"] == "B8"
    assert len(json.loads(request("GET", f"/api/documents/{doc['document_id']}/versions")[2])) == 1


def test_cong_thuc_goi_ra_ngoai_bi_tu_choi():
    status, body = save(new_doc(), [{"sheet": SHEET, "cell": "I8", "formula": 'WEBSERVICE("http://example.invalid")'}])
    assert status == 422 and body["code"] == "WRITE_REFUSED"


def test_chu_giong_cong_thuc_van_la_chu():
    doc = new_doc()
    save(doc, [{"sheet": SHEET, "cell": "I8", "value": "=1+1"}])
    cell = openpyxl.load_workbook(io.BytesIO(export(doc)))[SHEET]["I8"]
    assert cell.value == "=1+1" and cell.data_type != "f"


def test_thieu_nguoi_sua_bi_tu_choi():
    status, _ = post(f"/api/documents/{new_doc()['document_id']}/versions", {"base_version": 1, "author": "", "edits": [{"sheet": SHEET, "cell": "B8", "value": 1}]})
    assert status == 422


def test_project_rieng():
    assert PROJECT.endswith("test")
