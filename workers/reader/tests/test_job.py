"""Cổng đầu vào và bộ dịch, chạy không cần Docker."""
import io
import pathlib
import zipfile
from types import SimpleNamespace as N

import pytest

from reader import gate
from reader.job import process
from reader.ooxml import Colors, Package

FIX = pathlib.Path(__file__).resolve().parents[3] / "fixtures" / "templates"
QC = FIX / "GIA-QC-01_mau.xlsx"


def run(path: pathlib.Path, filename=None):
    return process(path.read_bytes(), filename or path.name, path)


def run_bytes(tmp_path, data: bytes, filename="thu.xlsx"):
    p = tmp_path / "in.xlsx"
    p.write_bytes(data)
    return process(data, filename, p)


def rezip(src: pathlib.Path, change=None, add=None) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(src) as zi, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for n in zi.namelist():
            d = zi.read(n)
            if change and n in change:
                d = change[n](d)
            zo.writestr(n, d)
        for n, d in (add or {}).items():
            zo.writestr(n, d)
    return out.getvalue()


def test_mau_qc_mo_duoc_va_bao_muc_chua_hien():
    r = run(QC)
    assert r["ok"]
    sheets = {s["name"]: s for s in r["grid"]["sheets"]}
    assert list(sheets) == ["KQ_QC_Thang", "GioiHan", "Bieu_do", "Huong_dan"]
    assert sheets["Huong_dan"]["hidden"] and not sheets["KQ_QC_Thang"]["hidden"]
    assert sheets["KQ_QC_Thang"]["merges"], "tiêu đề hai dòng có ô gộp"
    assert any(c.get("f") for c in sheets["KQ_QC_Thang"]["cells"])
    assert len(sheets["KQ_QC_Thang"]["cf"]) >= 6
    assert ("chart", "Bieu_do") in {(u["kind"], u["sheet"]) for u in r["unshown"]}
    assert any(u["kind"] == "comment" and u["sheet"] == "KQ_QC_Thang" and u["cell"] == "B7" for u in r["unshown"])
    assert r["blocked_formulas"] == [] and r["external"] == []


def test_kho_bao_shape_bang_thanh_du_lieu_va_hien_anh():
    r = run(FIX / "GIA-KHO-02_tieu_hao.xlsx")
    assert r["ok"]
    assert {"shape", "table_style", "data_bar"} <= {u["kind"] for u in r["unshown"]}
    first = r["grid"]["sheets"][0]
    assert first["images"] and first["images"][0]["src"].startswith("data:image/png;base64,")
    # Tham chiếu Table và tên vùng đã đổi sang A1, không còn dạng Bang[Cot].
    assert not any("[" in c.get("f", "") for s in r["grid"]["sheets"] for c in s["cells"])


def test_x18_cong_thuc_goi_ra_ngoai_khong_toi_luoi():
    r = run(FIX / "GIA-X18_ham_mang.xlsx")
    assert r["ok"]
    cells = {(b["sheet"], b["cell"]): b for b in r["blocked_formulas"]}
    assert cells[("KQ_QC_Thang", "I8")]["kind"] == "network"
    assert cells[("KQ_QC_Thang", "I9")]["kind"] == "external_link"
    text = str(r["grid"])
    assert "WEBSERVICE" not in text and "example.invalid" not in text and "ngoai_gia" not in text


def test_external_link_part_duoc_quet(tmp_path):
    """Liên kết ngoài thật nằm ở xl/externalLinks, không chỉ trong chữ công thức."""
    link = b'<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>'
    rel = (b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" '
           b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="http://example.invalid/x.xlsx" TargetMode="External"/></Relationships>')
    r = run_bytes(tmp_path, rezip(QC, add={"xl/externalLinks/externalLink1.xml": link, "xl/externalLinks/_rels/externalLink1.xml.rels": rel}))
    assert r["ok"]
    assert [e["kind"] for e in r["external"]] == ["external_link"]
    assert r["external"][0]["target"] == "http://example.invalid/x.xlsx"


@pytest.mark.parametrize("name,code", [
    ("GIA-X19_co_mat_khau.xlsx", "FILE_PASSWORD"),
    ("GIA-QC-01_da_dien_doi_cu.xls", "FILE_LEGACY_XLS"),
    ("GIA-X19_xls_mat_khau.xls", "FILE_PASSWORD"),
    ("GIA-X19_hong.xlsx", "FILE_CORRUPT"),
    ("GIA-QT-03_quy_trinh_doi_cu.doc", "FILE_UNSUPPORTED"),
])
def test_moi_loai_loi_co_ma_rieng(name, code):
    r = run(FIX / name)
    assert not r["ok"] and r["code"] == code


def test_xls_doi_duoi_thanh_xlsx_van_nhan_dung():
    assert run(FIX / "GIA-QC-01_da_dien_doi_cu.xls", "doi_ten.xlsx")["code"] == "FILE_LEGACY_XLS"


def test_mat_khau_doi_duoi_van_nhan_dung():
    assert run(FIX / "GIA-X19_co_mat_khau.xlsx", "doi_ten.xls")["code"] == "FILE_PASSWORD"


def test_txt_bao_dinh_dang(tmp_path):
    r = run_bytes(tmp_path, b"khong phai excel", "ghi_chu.txt")
    assert r["code"] == "FILE_UNSUPPORTED" and ".txt" in r["message"]


def test_zip_thuong_khong_phai_excel(tmp_path):
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as z:
        z.writestr("a.txt", "x")
    assert run_bytes(tmp_path, out.getvalue(), "a.zip")["code"] == "FILE_UNSUPPORTED"
    assert run_bytes(tmp_path, out.getvalue(), "a.xlsx")["code"] == "FILE_CORRUPT"


def test_xlsm_khong_chay_macro_va_bao():
    r = run(FIX / "GIA-X19_khong_co_vba.xlsm")
    assert r["ok"] and r["warnings"][0]["code"] == "FILE_MACRO_IGNORED" and not r["has_macro"]


def test_co_vbaproject_bao_co_macro(tmp_path):
    r = run_bytes(tmp_path, rezip(FIX / "GIA-X19_khong_co_vba.xlsm", add={"xl/vbaProject.bin": b"gia"}), "m.xlsm")
    assert r["ok"] and r["has_macro"] and "File này có phần tự động hóa" in r["warnings"][0]["message"]


def test_sheet_khoa_duoc_bao():
    assert run(FIX / "GIA-X19_sheet_khoa.xlsx")["protected_sheets"] == ["KQ_QC_Thang"]


def test_vuot_dung_luong_file(tmp_path):
    r = run_bytes(tmp_path, b"PK\x03\x04" + b"0" * (gate.MAX_FILE_BYTES + 1))
    assert r["code"] == "FILE_TOO_LARGE" and r["detail"]["limit"] == "file"


def test_bom_giai_nen(tmp_path):
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", "<Types/>")
        z.writestr("xl/workbook.xml", "<workbook/>")
        z.writestr("xl/bom.bin", b"\0" * (gate.MAX_UNCOMPRESSED_BYTES + 1024))
    data = out.getvalue()
    assert len(data) < gate.MAX_FILE_BYTES
    r = run_bytes(tmp_path, data)
    assert r["code"] == "FILE_TOO_LARGE" and r["detail"]["limit"] == "uncompressed"


def test_qua_nhieu_o(monkeypatch):
    monkeypatch.setattr(gate, "MAX_CELLS", 50)
    r = run(QC)
    assert r["code"] == "FILE_TOO_LARGE" and r["detail"]["limit"] == "cells"
    assert "ô dữ liệu, vượt giới hạn 200.000 ô" in r["message"]


@pytest.mark.parametrize("doctype", [
    b'<!DOCTYPE lolz [<!ENTITY a "aaaaaaaaaa"><!ENTITY b "&a;&a;&a;&a;">]>',
    b'<!DOCTYPE x SYSTEM "http://example.invalid/x.dtd">',
])
def test_xml_co_dtd_bi_chan(tmp_path, doctype):
    data = rezip(QC, change={"xl/sharedStrings.xml": lambda d: d.replace(b"<sst", doctype + b"<sst", 1)})
    r = run_bytes(tmp_path, data)
    assert r["code"] == "FILE_UNSAFE" and r["detail"]["part"] == "xl/sharedStrings.xml"


def test_ten_phan_nguy_hiem(tmp_path):
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as z:
        z.writestr("../../etc/passwd", "x")
    assert run_bytes(tmp_path, out.getvalue())["code"] == "FILE_CORRUPT"


def test_cong_thuc_bi_chan_trong_dinh_dang_co_dieu_kien(tmp_path):
    data = rezip(FIX / "GIA-QC-01_da_dien.xlsx", change={"xl/worksheets/sheet1.xml": lambda d: d.replace(
        b"<xm:f>AND(ISNUMBER(B8)", b'<xm:f>AND(WEBSERVICE("http://example.invalid/cf")&lt;&gt;"",ISNUMBER(B8)', 1)})
    r = run_bytes(tmp_path, data)
    assert r["ok"]
    assert any(b["where"] == "conditional_format" for b in r["blocked_formulas"])
    assert "example.invalid" not in str(r["grid"])


def test_mau_theme_duoc_doi():
    with zipfile.ZipFile(QC) as z:
        c = Colors(Package(z))
    assert c.resolve(N(type="theme", theme=1, tint=0, rgb=None, indexed=None)) == "#000000"
    assert c.resolve(N(type="theme", theme=0, tint=-0.5, rgb=None, indexed=None)) == "#808080"


BANNED = ("BadZip", "zip", "XML", "encrypted", "exception", "error", "FILE_", "HTTP", "VBA", "Univer", "worker", "backend", "upload", "shape", "Table")


@pytest.mark.parametrize("name", ["GIA-X19_co_mat_khau.xlsx", "GIA-QC-01_da_dien_doi_cu.xls", "GIA-X19_hong.xlsx", "GIA-QT-03_quy_trinh_doi_cu.doc"])
def test_thong_bao_doc_duoc(name):
    """Theo acceptance_m1.md mục 7: không có từ kỹ thuật, tối đa 3 câu, nói file có bị đổi không."""
    msg = run(FIX / name)["message"]
    assert not [w for w in BANNED if w.lower() in msg.lower()]
    assert "không bị đổi" in msg
    assert msg.count(". ") <= 2
