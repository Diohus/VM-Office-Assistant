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
