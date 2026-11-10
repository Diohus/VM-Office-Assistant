"""Cổng đầu vào qua API thật, kho file, xuất lại nguyên bản, cách ly mạng."""
from __future__ import annotations

import hashlib
import subprocess

import pytest

from conftest import BASE, FIX, GENERATED, PROJECT, request, upload

SAMPLES = ["GIA-QC-01_mau.xlsx", "GIA-KHO-02_tieu_hao.xlsx", "GIA-X19_sheet_khoa.xlsx"]


@pytest.mark.parametrize("name", SAMPLES)
def test_xuat_lai_giong_het_tung_byte_va_ten_khac(name):
    data = (FIX / name).read_bytes()
    status, doc = upload(data, name)
    assert status == 201
    assert doc["sha256"] == hashlib.sha256(data).hexdigest()
    status, headers, body = request("GET", f"/api/documents/{doc['document_id']}/export")
    assert status == 200 and body == data
    disposition = {k.lower(): v for k, v in headers.items()}["content-disposition"]
    assert "_xuat.xlsx" in disposition and f"filename*=UTF-8''{name}" not in disposition


@pytest.mark.parametrize("name,code,status", [
    ("GIA-X19_co_mat_khau.xlsx", "FILE_PASSWORD", 422),
    ("GIA-X19_hong.xlsx", "FILE_CORRUPT", 422),
    ("GIA-X19_xls_mat_khau.xls", "FILE_PASSWORD", 422),
])
def test_loi_tung_loai(name, code, status):
    got, body = upload((FIX / name).read_bytes(), name)
    assert got == status and body["code"] == code


@pytest.mark.parametrize("name,limit", [
    ("GIA-VUOT_25MB.xlsx", "file"),
    ("GIA-VUOT_giai_nen_200MB.xlsx", "uncompressed"),
    ("GIA-VUOT_200000_o.xlsx", "cells"),
])
def test_moi_gioi_han_bao_dung_loai(name, limit):
    got, body = upload((GENERATED / name).read_bytes(), name)
    assert got == 413 and body["code"] == "FILE_TOO_LARGE" and body["detail"]["limit"] == limit


def test_file_rat_lon_bi_chan_o_cua():
    got, body = upload(b"\0" * (27 * 1024 * 1024), "lon.xlsx")
    assert got == 413 and body["code"] == "FILE_TOO_LARGE" and "25 MB" in body["message"]


def test_txt_doi_duoi_xlsx():
    got, body = upload(b"khong phai excel", "gia.xlsx")
    assert got == 422 and body["code"] == "FILE_CORRUPT"


def test_ten_file_la_khong_gay_hai():
    got, doc = upload((FIX / "GIA-QC-01_mau.xlsx").read_bytes(), "..\\..\\etc\\<b>ten\x07.xlsx")
    assert got == 201 and doc["name"] == "<b>ten.xlsx"


def test_id_la_tra_404():
    assert request("GET", "/api/documents/khong-phai-uuid")[0] == 404
    assert request("GET", "/api/documents/00000000-0000-0000-0000-000000000000/grid")[0] == 404


def test_x18_bao_o_lien_quan():
    _, doc = upload((FIX / "GIA-X18_ham_mang.xlsx").read_bytes(), "x18.xlsx")
    assert {b["cell"] for b in doc["blocked_formulas"]} == {"I8", "I9"}


def test_x18b_lien_ket_ngoai_that():
    _, doc = upload((FIX / "GIA-X18b_lien_ket_ngoai.xlsx").read_bytes(), "x18b.xlsx")
    assert {b["cell"] for b in doc["blocked_formulas"]} == {"I8", "I9"}, "I9 chỉ lộ liên kết qua tên vùng"
    assert {e["kind"] for e in doc["external"]} == {"external_link"}, "Excel khai 2 đường dẫn cho một liên kết"
    assert len({e["part"] for e in doc["external"]}) == 1


def _exec(service: str, code: str) -> subprocess.CompletedProcess:
    return subprocess.run(["docker", "exec", f"{PROJECT}-{service}-1", "python", "-c", code], capture_output=True, text=True, timeout=60)


REACH_OUT = "import socket; socket.create_connection(('1.1.1.1', 443), 3); print('REACHED')"


@pytest.mark.parametrize("service", ["reader", "api"])
def test_khong_co_duong_ra_ngoai(service):
    proc = _exec(service, REACH_OUT)
    assert "REACHED" not in proc.stdout and proc.returncode != 0


def test_reader_khong_noi_duoc_db():
    proc = _exec("reader", "import socket; socket.create_connection(('db', 5432), 3); print('REACHED')")
    assert "REACHED" not in proc.stdout


@pytest.mark.parametrize("service", ["reader", "api"])
def test_khong_co_khoa_va_chi_doc(service):
    # Chỉ in tên biến, không in giá trị. GPG_KEY là khóa công khai của image Python.
    proc = _exec(service, "import os; print([k for k in os.environ if any(w in k.upper() for w in ('API', 'KEY', 'TOKEN')) and k != 'GPG_KEY'])")
    assert proc.stdout.strip() == "[]"
    proc = _exec(service, "open('/opt/app/x', 'w')")
    assert proc.returncode != 0


def test_reader_khong_thay_mat_khau_db():
    proc = _exec("reader", "import os; print('DATABASE_URL' in os.environ, 'POSTGRES_PASSWORD' in os.environ)")
    assert proc.stdout.strip() == "False False"


def test_spool_khong_con_file_sau_khi_xu_ly():
    upload((FIX / "GIA-X19_hong.xlsx").read_bytes(), "a.xlsx")
    upload((FIX / "GIA-QC-01_mau.xlsx").read_bytes(), "b.xlsx")
    proc = _exec("api", "import os; print(os.listdir('/spool'))")
    assert proc.stdout.strip() == "[]"


def test_file_bi_tu_choi_khong_vao_kho():
    data = (FIX / "GIA-X19_hong.xlsx").read_bytes()
    upload(data, "a.xlsx")
    sha = hashlib.sha256(data).hexdigest()
    proc = _exec("api", f"import os; print(os.path.exists('/store/{sha[:2]}/{sha}.xlsx'))")
    assert proc.stdout.strip() == "False"


def test_base_la_cuc_bo():
    assert BASE.startswith("http://127.0.0.1")
