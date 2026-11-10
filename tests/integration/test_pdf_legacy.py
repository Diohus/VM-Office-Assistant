"""Xuất PDF (Q48c), file đời cũ (Q57, Q62), worker LibreOffice không mạng, qua API thật."""
from __future__ import annotations

import io
import json
import subprocess
import time
import zipfile

import openpyxl
import pdfplumber
from openpyxl.utils.datetime import to_excel
import pytest

from conftest import FIX, PROJECT, request, upload

SHEET = "KQ_QC_Thang"
JSON = {"Content-Type": "application/json"}


def post(path: str, body: dict):
    return request("POST", path, json.dumps(body).encode("utf8"), headers=JSON)


def doc_qc() -> dict:
    return upload((FIX / "GIA-QC-01_da_dien.xlsx").read_bytes(), "GIA-QC-01_da_dien.xlsx")[1]


def cached_values() -> list[dict]:
    """Số Excel đã lưu trong fixture đóng vai số trên lưới: 239/239 đã khớp lưới ở test trình duyệt. Ngày gửi dạng số như lưới gửi thật."""
    formulas = openpyxl.load_workbook(FIX / "GIA-QC-01_da_dien.xlsx")
    values = openpyxl.load_workbook(FIX / "GIA-QC-01_da_dien.xlsx", data_only=True)
    out = []
    for ws in formulas.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.data_type == "f":
                    v = values[ws.title][c.coordinate].value
                    out.append({"sheet": ws.title, "cell": c.coordinate, "value": to_excel(v) if hasattr(v, "isoformat") else v})
    return out


def test_pdf_khop_luoi_co_tieng_viet_va_so_doc_ra_duoc():
    doc = doc_qc()
    status, headers, data = post(f"/api/documents/{doc['document_id']}/pdf", {"grid_values": cached_values()})
    assert status == 200, data[:400]
    assert data[:5] == b"%PDF-"
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        text = pdf.pages[0].extract_text()
    assert "MẪU GIẢ" in text and "NHẬT KÝ" in text, "chữ có dấu phải đúng"
    print("giây chuyển đổi:", {k.lower(): v for k, v in headers.items()}["x-convert-seconds"])


def test_pdf_giu_kho_giay_a4_ca_khi_sheet_khong_khai():
    doc = doc_qc()
    status, _, data = post(f"/api/documents/{doc['document_id']}/pdf", {"grid_values": cached_values()})
    assert status == 200
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        sizes = {tuple(sorted((round(p.width), round(p.height)))) for p in pdf.pages}
        assert len(pdf.pages) == 8, "số trang như Excel"
    assert sizes == {(595, 842)}, "A4, không phải Letter"


def test_pdf_lech_thi_bao_o_nao_khong_xuat():
    doc = doc_qc()
    values = cached_values()
    target = next(v for v in values if v["cell"] == "B39")
    target["value"] = 999.0
    status, _, data = post(f"/api/documents/{doc['document_id']}/pdf", {"grid_values": values})
    body = json.loads(data)
    assert status == 409 and body["code"] == "PDF_MISMATCH"
    assert [c["cell"] for c in body["detail"]["cells"]] == ["B39"] and "chưa xuất PDF" in body["message"]


def test_pdf_ban_da_sua_lo_tinh_lai_khi_khong_co_so_da_luu():
    doc = doc_qc()
    post(f"/api/documents/{doc['document_id']}/versions", {"base_version": 1, "author": "Người Giả", "edits": [{"sheet": SHEET, "cell": "B8", "value": 500}]})
    status, _, data = post(f"/api/documents/{doc['document_id']}/pdf", {"grid_values": []})
    assert status == 200 and data[:5] == b"%PDF-"
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        text = "\n".join(p.extract_text() or "" for p in pdf.pages)
    assert "22.01" in text and "22,01" not in text, "B39 sau khi sửa B8 thành 500, dấu chấm như lưới (Q122)"


def test_xls_mo_duoc_co_canh_bao_va_giu_ban_goc():
    data = (FIX / "GIA-QC-01_da_dien_doi_cu.xls").read_bytes()
    status, doc = upload(data, "GIA-QC-01_da_dien_doi_cu.xls")
    assert status == 201, doc
    assert doc["name"] == "GIA-QC-01_da_dien_doi_cu.xlsx"
    assert doc["warnings"][0]["code"] == "FILE_LEGACY_CONVERTED" and "Bản gốc còn nguyên" in doc["warnings"][0]["message"]
    exported = request("GET", f"/api/documents/{doc['document_id']}/export")[2]
    with zipfile.ZipFile(io.BytesIO(exported)) as z:
        assert b"_xlfn_IFERROR" not in z.read("xl/workbook.xml"), "tên vùng rỗng làm Excel không mở được"
    assert openpyxl.load_workbook(io.BytesIO(exported))[SHEET]["B39"].value


def test_docker_worker_khong_co_mang_va_chi_doc():
    run = lambda code: subprocess.run(["docker", "exec", f"{PROJECT}-converter-1", "python3", "-c", code], capture_output=True, text=True, timeout=60)
    assert "REACHED" not in run("import socket; socket.create_connection(('1.1.1.1', 443), 3); print('REACHED')").stdout
    assert run("import os; print(os.listdir('/sys/class/net'))").stdout.strip() == "['lo']"
    assert run("open('/opt/x', 'w')").returncode != 0


@pytest.fixture
def converter_stopped():
    name = f"{PROJECT}-converter-1"
    subprocess.run(["docker", "stop", "-t", "1", name], check=True, capture_output=True, timeout=60)
    time.sleep(11)  # hơn thời gian API coi worker là còn sống
    try:
        yield
    finally:
        subprocess.run(["docker", "start", name], check=True, capture_output=True, timeout=60)
        time.sleep(3)


def test_tat_worker_bao_ro_va_excel_van_xuat_duoc(converter_stopped):
    doc = doc_qc()
    started = time.monotonic()
    status, _, data = post(f"/api/documents/{doc['document_id']}/pdf", {"grid_values": []})
    assert status == 503 and "file Excel" in json.loads(data)["message"]
    assert time.monotonic() - started < 5, "báo ngay, không chờ hết giờ"
    assert request("GET", f"/api/documents/{doc['document_id']}/export")[0] == 200
    status, body = upload((FIX / "GIA-QC-01_da_dien_doi_cu.xls").read_bytes(), "a.xls")
    assert status == 422 and body["code"] == "FILE_LEGACY_FAILED" and "Bản gốc không bị đổi" in body["message"]
