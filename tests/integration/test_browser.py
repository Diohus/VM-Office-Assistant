"""Luồng thật trên trình duyệt: chạy browser/run.mjs bằng Edge có sẵn, rồi kiểm tra kết quả."""
from __future__ import annotations

import hashlib
import io
import json
import os
import pathlib
import subprocess
import urllib.request
import zipfile

import pytest

from conftest import BASE, FIX, ROOT

HERE = pathlib.Path(__file__).parent
OUT = HERE / "out"


def canary_copy(url: str) -> pathlib.Path:
    """Bản sao X18 trỏ địa chỉ tới máy chủ bẫy, thêm liên kết ngoài và ảnh từ xa ở cấp gói."""
    OUT.mkdir(exist_ok=True)
    out = OUT / "x18_canary.xlsx"
    rels = ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="{url}/link.xlsx" TargetMode="External"/>'
            '</Relationships>')
    with zipfile.ZipFile(FIX / "GIA-X18_ham_mang.xlsx") as zi, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for n in zi.namelist():
            zo.writestr(n, zi.read(n).replace(b"http://example.invalid/gia", f"{url}/webservice".encode()))
        zo.writestr("xl/externalLinks/externalLink1.xml", '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>')
        zo.writestr("xl/externalLinks/_rels/externalLink1.xml.rels", rels)
    return out


@pytest.fixture(scope="module")
def run(canary):
    copy = canary_copy(canary["url"])
    # Đối chứng: máy chủ bẫy phải ghi nhận được, nếu không thì "0 lần gọi" chẳng có nghĩa gì.
    urllib.request.urlopen(canary["url"] + "/doi-chung", timeout=5).read()
    assert canary["hits"] == ["/doi-chung"]
    canary["hits"].clear()
    node = "node.exe" if os.name == "nt" else "node"
    proc = subprocess.run([node, "run.mjs"], cwd=HERE / "browser", capture_output=True, text=True, encoding="utf8", timeout=570,
                          env={**os.environ, "VM_BASE": BASE, "VM_CANARY_FILE": str(copy), "VM_CANARY_URL": canary["url"]})
    assert proc.returncode == 0, proc.stderr[-2000:]
    data = json.loads((OUT / "browser.json").read_text(encoding="utf8"))
    data["canary_hits"] = list(canary["hits"])
    return data


def test_khong_co_yeu_cau_ra_ngoai_tren_moi_file(run):
    assert run["outbound"] == []
    assert run["canary_hits"] == []


def test_khong_loi_trang_va_khong_vi_pham_csp(run):
    assert run["page_errors"] == []
    assert run["csp_violations"] == []


def test_mau_qc(run):
    r = run["files"]["GIA-QC-01_mau.xlsx"]
    assert r["sheets"] == ["KQ_QC_Thang", "GioiHan", "Bieu_do"], "Huong_dan là sheet ẩn"
    assert r["problems"] == []
    assert r["mismatches"] == []
    text = " ".join(n["text"] for n in r["notices"])
    assert "1 biểu đồ" in text and "1 nhận xét" in text and "còn nguyên" in text
    assert "Huong_dan" not in json.dumps(r["notices"], ensure_ascii=False)


def test_xem_vi_tri_dua_toi_dung_o(run):
    assert run["files"]["GIA-QC-01_mau.xlsx"]["jump_comment"] == {"sheet": "KQ_QC_Thang", "cell": "B7"}
    assert run["files"]["GIA-QC-01_mau.xlsx"]["jump_chart"]["sheet"] == "Bieu_do"


def test_nhan_xet_doc_duoc_trong_app(run):
    assert "Mức 1 là mức thấp" in run["files"]["GIA-QC-01_mau.xlsx"]["comment_text"]


def test_thong_bao_dong_duoc_va_xem_lai_duoc(run):
    r = run["files"]["GIA-QC-01_mau.xlsx"]
    assert r["after_close"] == 0 and r["after_reopen"] >= 1


def test_mau_da_dien_cong_thuc_khop_excel(run):
    r = run["files"]["GIA-QC-01_da_dien.xlsx"]
    assert r["formulas"] >= 150 and r["mismatches"] == [] and r["problems"] == []
    assert abs(r["B39"] - 5.5256666666) < 1e-6
    assert r["cf_rules"] >= 6
    assert r["dv_H8"] == "list"


def test_kho(run):
    r = run["files"]["GIA-KHO-02_tieu_hao.xlsx"]
    assert r["formulas"] >= 80 and r["problems"] == []
    assert r["mismatches"] == [], "CELL, INFO bị chặn khi mở nên chỉ hiện số đã lưu"
    text = " ".join(n["text"] for n in r["notices"])
    assert "hình vẽ" in text and "kiểu trình bày của bảng" in text and "thanh dữ liệu" in text and "ảnh" not in text.replace("trình bày", "")


def test_nguon_khong_co_thong_bao(run):
    assert run["files"]["GIA-NGUON_2026-09.xlsx"]["notices"] == []


def test_x18_canh_bao(run):
    text = " ".join(n["text"] for n in run["files"]["x18_canary.xlsx"]["notices"])
    assert "2 ô lấy dữ liệu từ ngoài file" in text and "App không gọi ra ngoài" in text and "1 liên kết" in text


def test_loi_roi_mo_file_dung_khong_can_tai_lai_trang(run):
    seq = run["error_sequence"]
    for name in ("GIA-X19_co_mat_khau.xlsx", "GIA-X19_hong.xlsx", "ghi_chu.txt"):
        assert "không bị đổi" in seq[name]
    assert len(set(seq.values())) == 3, "mỗi loại lỗi có lời riêng"
    assert "mật khẩu" in seq["GIA-X19_co_mat_khau.xlsx"]
    assert run["recovered_doc"] == "GIA-QC-01_mau.xlsx"


def test_mat_ket_noi_bao_duoc(run):
    assert "mất kết nối" in run["network_down"] and "không bị đổi" in run["network_down"]


def test_tat_api_that_bao_duoc_va_tu_hoi_phuc(run):
    assert "mất kết nối" in run["backend_down"] and "không bị đổi" in run["backend_down"]
    assert run["after_backend_back"] == "GIA-QC-01_mau.xlsx"


def test_moi_o_cung_mot_kieu_dau_thap_phan(run):
    assert run["files"]["GIA-X18b_lien_ket_ngoai.xlsx"]["display_general"] == ["1.23", "4.56", "5.51"]


def test_o_general_cot_hep_lam_tron_vua_cot_nhu_excel(run):
    r = run["files"]["GIA-KHO-02_tieu_hao.xlsx"]
    assert r["display_narrow"] == {"C5": "5.969192", "C6": "5.969192", "C7": "5.662873", "C15": "16.08333"}
    assert r["values_narrow"]["C5"] != 5.969192, "chỉ làm tròn khi hiện, giá trị giữ nguyên"
    assert [c for c in r["display_all"] if "#" in c["text"].replace("#N/A", "").replace("#REF!", "").replace("#DIV/0!", "").replace("#NAME?", "").replace("#VALUE!", "")] == []


def test_khong_lo_sheet_an_qua_nut_danh_sach(run):
    r = run["files"]["GIA-X18b_lien_ket_ngoai.xlsx"]
    assert r["sheet_list_button"] == 0
    assert "Huong_dan" not in r["body_text"]


def test_mau_theo_ui_spec(run):
    assert run["files"]["GIA-X18b_lien_ket_ngoai.xlsx"]["btn_color"] == "#1b3a6b"
    assert "0a5fa8" not in (ROOT / "frontend/src/theme.css").read_text(encoding="utf8").lower()


def test_nut_xuat_tai_dung_file(run):
    assert run["download"]["sha256"] == hashlib.sha256((FIX / "GIA-QC-01_mau.xlsx").read_bytes()).hexdigest()
    assert run["download"]["name"] == "GIA-QC-01_mau_xuat.xlsx"


def test_bo_cuc_1366(run):
    lay = run["layout"]
    assert lay["page_scroll_x"] is False
    assert 44 <= lay["topbar_h"] <= 52 and 68 <= lay["sidebar_w"] <= 76
    assert lay["grid_share"] > 0.6
    assert lay["sidebar_labels"] == ["Mở", "Lưu", "Xuất", "Điền từ file", "Chat AI", "Phiên bản"]  # lát 1 thêm Điền từ file (Q133), lát 2 Chat AI (Q135)


def test_n9_ham_thong_ke_vung_it_so_khop_excel(run):
    # R-1015: STDEV, VAR (và S, A) vùng không có số; MINA, MAXA vùng có ô trống. Đáp án là số Excel thật đã lưu.
    r = run["files"]["GIA-X20_thong_ke_vung_it_so.xlsx"]
    assert r["formulas"] >= 240 and r["problems"] == []
    assert r["mismatches"] == []


def test_n9_pdf_bieu_mau_qc_trong_xuat_duoc(run):
    p = run["blank_pdf"]
    assert p["sd_B40"] in ("", None), "SD của cột chưa có số phải trống như Excel, không phải 0"
    assert p["error"] is None and p["downloaded"] is True and p["name"] == "GIA-QC-01_mau_xuat.pdf"


def test_lat2_if_round_abs_and_or_khop_excel(run):
    # Q135: hàm cho phép của khung chat phải được lưới tính như Excel. Có số khó làm tròn (2,675; 1,005; -2,5), trống, chữ, TRUE.
    r = run["files"]["GIA-X21_ham_logic_lam_tron.xlsx"]
    assert r["formulas"] >= 400 and r["problems"] == []
    assert r["mismatches"] == []
