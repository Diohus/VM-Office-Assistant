"""Q73: không có đường dẫn máy, tên người dùng, khóa trong những gì sẽ commit, kể cả bên trong file Office."""
from __future__ import annotations

import io
import pathlib
import sys
import zipfile

import pytest

ROOT = pathlib.Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "tools" / "commit_scan"))

import scan  # noqa: E402

FIX = ROOT / "fixtures" / "templates"
FAKE_KEY = "sk-" + "a" * 30


def with_part(src: pathlib.Path, part: str, payload: bytes) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(src) as zi, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for n in zi.namelist():
            zo.writestr(n, zi.read(n))
        zo.writestr(part, payload)
    return out.getvalue()


def test_danh_sach_commit_sach():
    files = scan.commit_files()
    assert files, "danh sách commit rỗng"
    bad = []
    for path in files:
        if path in scan.TEXT_ALLOW:
            continue
        bad += [f"{path} {part} {kind}" for part, kind in scan.scan_file((ROOT / path).read_bytes(), path)]
    assert bad == []


def test_danh_sach_khong_gom_tai_lieu_noi_bo_va_bi_mat():
    files = scan.commit_files()
    for path in files:
        assert not path.startswith(("docs/", "spikes/", ".local/")), path
        assert path not in ("chat_log.md", "customer_versions.txt", "output_natural.txt", "project_requirement.txt", ".env"), path
        assert "~$" not in path and not path.endswith(".pem"), path
    assert "fixtures/templates/GIA-KHO-02_kho.xlsx" not in files


@pytest.mark.parametrize("payload,kind", [
    (rb'<x15ac:absPath url="D:\Vin\VM-Office-Assistant\fixtures\templates\"/>', "duong_dan_o_dia"),
    (rb'<Relationship Target="file:///X:/a/b.xlsx"/>', "duong_dan_file"),
    (rb'<a>\\may-chu\chia-se\file.xlsx</a>', "duong_dan_mang"),
    (rb'<a>C:\Users\ai_do\x</a>', "thu_muc_nguoi_dung"),
    (FAKE_KEY.encode(), "khoa_openai_anthropic"),
])
def test_phep_quet_truot_khi_co_trong_file_office(payload, kind):
    data = with_part(FIX / "GIA-QC-01_mau.xlsx", "xl/them.xml", payload)
    assert ("xl/them.xml", kind) in scan.scan_file(data, "x.xlsx", extra_terms=[])


def test_phep_quet_bat_ban_chua_don():
    """Đúng ca đã xảy ra: workbook.xml có absPath do Excel ghi."""
    data = with_part(FIX / "GIA-QC-01_mau.xlsx", "xl/workbook2.xml", rb'<mc:Choice><x15ac:absPath url="D:\Du_an\"/></mc:Choice>')
    assert any(kind == "duong_dan_o_dia" for _, kind in scan.scan_file(data, "x.xlsx", extra_terms=[]))


def test_phep_quet_bat_ten_nguoi_dung_dang_utf16_trong_file_doi_cu():
    data = b"\xd0\xcf\x11\xe0" + "tac gia: tenriengxyz".encode("utf-16-le")
    assert ("", "ten_nguoi_dung") in scan.scan_file(data, "x.doc", extra_terms=[b"tenriengxyz"])


def test_khong_bao_nham_bang_excel_binh_thuong():
    data = with_part(FIX / "GIA-QC-01_mau.xlsx", "xl/them.xml", b"<a>Mau gia: A1:B2, 5:30, http://schemas.openxmlformats.org/x</a>")
    assert [k for p, k in scan.scan_file(data, "x.xlsx", extra_terms=[]) if p == "xl/them.xml"] == []


# ---------- R-1016: tên người trong thông tin file Office ----------

import people_meta  # noqa: E402
import olefile  # noqa: E402

FAKE_NAME = "Nguoi Gia Thu"


def with_core(src: pathlib.Path, creator: str, last: str, company: str | None = None) -> bytes:
    """Bản sao của src với core.xml (và app.xml nếu có company) mang giá trị cho trước."""
    core = ('<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/">'
            f"<dc:creator>{creator}</dc:creator><cp:lastModifiedBy>{last}</cp:lastModifiedBy></cp:coreProperties>").encode()
    out = io.BytesIO()
    with zipfile.ZipFile(src) as zi, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for n in zi.namelist():
            d = zi.read(n)
            if n == "docProps/core.xml":
                d = core
            elif n == "docProps/app.xml" and company is not None:
                d = d.replace(b"</Properties>", f"<Company>{company}</Company></Properties>".encode())
            zo.writestr(n, d)
    return out.getvalue()


def ole_with_name(src: pathlib.Path, tmp_path: pathlib.Path, stream: str) -> bytes:
    """Ghi tên giả vào các ô tên của một luồng trong file đời cũ đã sạch (đúng chỗ Excel/Word ghi)."""
    data = src.read_bytes()
    with olefile.OleFileIO(io.BytesIO(data)) as o:
        fields = [f for f in people_meta._ole_fields(o) if f[0] == stream]
        body = bytearray(o.openstream(stream).read())
    assert fields, f"{src.name} không có ô tên trong {stream}"
    for _, _, at, ln, _, fill in fields:
        text = FAKE_NAME.encode("utf-16-le" if fill == b" \x00" else "latin-1")[:ln]
        body[at: at + len(text)] = text
    path = tmp_path / ("x" + src.suffix)
    path.write_bytes(data)
    with olefile.OleFileIO(str(path), write_mode=True) as o:
        o.write_stream(stream, bytes(body))
    return path.read_bytes()


def people_findings(data: bytes, name: str) -> list[tuple[str, str]]:
    return [(p, k) for p, k in scan.scan_file(data, name, extra_terms=[]) if k in ("ten_trong_thong_tin_file", "thong_tin_file_ma_hoa", "khong_doc_duoc_thong_tin_file")]


@pytest.mark.parametrize("creator,last,company,where", [
    (FAKE_NAME, "", None, "docProps/core.xml dc:creator"),
    ("openpyxl", FAKE_NAME, None, "docProps/core.xml cp:lastModifiedBy"),
    ("", "", FAKE_NAME, "docProps/app.xml Company"),
])
def test_bat_ten_nguoi_trong_docprops(creator, last, company, where):
    data = with_core(FIX / "GIA-QC-01_mau.xlsx", creator, last, company)
    assert people_findings(data, "x.xlsx") == [(where, "ten_trong_thong_tin_file")]
    cleaned, n = people_meta.strip(data)
    assert n == 1 and people_findings(cleaned, "x.xlsx") == []


@pytest.mark.parametrize("value", ["", "openpyxl", "VM-Office-Assistant", "VM-Office-Assistant [MẪU GIẢ]"])
def test_gia_tri_cho_phep_khong_bi_bao_va_khong_bi_go(value):
    data = with_core(FIX / "GIA-QC-01_mau.xlsx", value, value)
    assert people_findings(data, "x.xlsx") == []
    assert people_meta.strip(data)[1] == 0


def test_go_ten_giu_nguyen_cac_phan_khac_cua_zip():
    data = with_core(FIX / "GIA-QC-01_mau.xlsx", FAKE_NAME, FAKE_NAME)
    cleaned, n = people_meta.strip(data)
    assert n == 2
    with zipfile.ZipFile(io.BytesIO(data)) as a, zipfile.ZipFile(io.BytesIO(cleaned)) as b:
        assert a.namelist() == b.namelist()
        assert [p for p in a.namelist() if a.read(p) != b.read(p)] == ["docProps/core.xml"]
        assert FAKE_NAME.encode() not in b.read("docProps/core.xml")


@pytest.mark.parametrize("src,stream,where", [
    ("GIA-QC-01_da_dien_doi_cu.xls", "\x05SummaryInformation", "SummaryInformation nguoi_luu_cuoi"),
    ("GIA-QC-01_da_dien_doi_cu.xls", "Workbook", "Workbook nguoi_luu_writeaccess"),
    ("GIA-QT-03_quy_trinh_doi_cu.doc", "\x05SummaryInformation", "SummaryInformation tac_gia"),
    ("GIA-QT-03_quy_trinh_doi_cu.doc", "1Table", "1Table tac_gia"),
])
def test_bat_va_go_ten_trong_file_doi_cu(tmp_path, src, stream, where):
    data = ole_with_name(FIX / src, tmp_path, stream)
    assert (where, "ten_trong_thong_tin_file") in people_findings(data, src)
    cleaned, n = people_meta.strip(data)
    assert n >= 1 and people_findings(cleaned, src) == []
    with olefile.OleFileIO(io.BytesIO(data)) as a, olefile.OleFileIO(io.BytesIO(cleaned)) as b:
        changed = ["/".join(s) for s in a.listdir() if a.openstream(s).read() != b.openstream(s).read()]
    assert changed == [stream]


def test_file_ma_hoa_ngoai_danh_sach_bi_bao():
    data = (FIX / "GIA-X19_co_mat_khau.xlsx").read_bytes()
    assert people_meta.kind(data) == "ole_ma_hoa"
    assert people_findings(data, "fixtures/templates/GIA-X19_co_mat_khau.xlsx") == []
    assert people_findings(data, "fixtures/templates/khac.xlsx") == [("", "thong_tin_file_ma_hoa")]
    xls = (FIX / "GIA-X19_xls_mat_khau.xls").read_bytes()
    assert people_meta.kind(xls) == "ole_ma_hoa"
    assert people_findings(xls, "fixtures/templates/khac.xls") == [("", "thong_tin_file_ma_hoa")]


def test_file_mau_khong_con_ten_nguoi():
    """Mọi file Office sẽ commit: không còn tên người ở thông tin file (R-1016)."""
    bad = []
    for path in scan.commit_files():
        bad += [f"{path} {p}" for p, _ in people_findings((ROOT / path).read_bytes(), path)]
    assert bad == []
