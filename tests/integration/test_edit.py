"""Luồng sửa ô, lưu, phiên bản trên trình duyệt thật. Số so với Excel thật nằm ở tools/excel_check/q48_check.py."""
from __future__ import annotations

import json
import os
import pathlib
import subprocess

import pytest

from conftest import BASE

HERE = pathlib.Path(__file__).parent
NEW_B39 = (500 - 5.51) / 30  # B8 từ 5,51 thành 500, cột có 30 số


@pytest.fixture(scope="module")
def run():
    node = "node.exe" if os.name == "nt" else "node"
    proc = subprocess.run([node, "edit.mjs"], cwd=HERE / "browser", capture_output=True, text=True, encoding="utf8", timeout=400, env={**os.environ, "VM_BASE": BASE})
    assert proc.returncode == 0, proc.stderr[-2500:]
    return json.loads((HERE / "out" / "edit.json").read_text(encoding="utf8"))


def test_khong_loi_va_khong_ra_ngoai(run):
    assert run["outbound"] == [] and run["page_errors"] == [] and run["csp_violations"] == []


def test_sua_so_cong_thuc_tinh_lai_ngay(run):
    assert run["typed"]["comma"]["value"] == 5.5 and run["typed"]["comma"]["display"] == "5.50"
    assert run["b39_after"] != run["b39_before"]
    assert run["formula_b39"].startswith("=")


def test_go_dau_cham_va_ca_hai_dau(run):
    assert run["typed"]["dot"] == 5.51 and run["typed"]["dot_display"] == "5.51", "giá trị đúng mà hiện sai cũng là lỗi"
    assert run["typed"]["both"] == 1234.5


def test_so_nhap_nhang_khong_doan(run):
    amb = run["typed"]["ambiguous_dot"]
    assert amb["value"] == amb["before"], "chưa được ghi"
    assert "số nghìn 1234" in amb["notice"] and "số lẻ 1.234" in amb["notice"] and "1.2340" in amb["notice"] and "không bị đổi" in amb["notice"]
    assert run["typed"]["ambiguous_comma"]["notice"] is not None


def test_go_dau_phay_la_so_le_co_bao_khong_thanh_so_nghin(run):
    # Q122, dòng chặn: app dùng dấu chấm; 5,5 không bao giờ lặng lẽ thành 55.
    comma = run["typed"]["comma"]
    assert comma["value"] == 5.5 and '"5,5" là 5.5' in comma["notice"] and "dấu chấm" in comma["notice"]
    swapped = run["typed"]["swapped"]
    assert swapped["0,5"]["value"] == 0.5 and '"0,5" là 0.5' in swapped["0,5"]["notice"]
    assert swapped["12,5"]["value"] == 12.5 and '"12,5" là 12.5' in swapped["12,5"]["notice"]
    unclear = run["typed"]["unclear"]
    assert unclear["value"] == unclear["before"] and "không phải một số rõ ràng" in unclear["notice"] and "dấu chấm" in unclear["notice"]


def test_chu_va_cong_thuc(run):
    assert run["typed"]["text"] == "abc"
    assert run["typed"]["formula_text"] == "=SUM(B8:B10)" and isinstance(run["typed"]["formula_value"], (int, float))


def test_dau_cham_giu_sau_khi_tinh_lai(run):
    # Q122: ô hiện dấu chấm, theo frontend/src/numberFormat.js. Gõ 1,5 vẫn là 1.5.
    assert run["general_comma"] == "1.5"
    assert run["general_formula_display"].startswith("0.3")


def test_hoan_tac_lam_lai(run):
    assert run["after_undo"] != run["undo_expected"] and run["after_redo"] == run["undo_expected"]


def test_q86_khong_mat_im_lang(run):
    assert run["dirty_count"] > 0 and run["after_stay_dirty"] == run["dirty_count"]
    assert "chưa lưu" in run["unsaved_dialog"]
    assert run["beforeunload_dialogs"] == 1


def test_luu_tao_ban_moi_va_khong_bao_lech_sai(run):
    assert run["v2_state"] == "Bản 2" and "Đã lưu bản 2 trong app, ghi 1 ô" in run["saved_notice"] and "File trong folder" in run["saved_notice"] and "bấm Xuất" in run["saved_notice"]
    assert run["v2_b8"] == 500 and run["v2_dirty"] == 0
    assert run["v2_mismatch_notice"] is False
    assert abs(run["v2_b39"] - (run["v1_b39"] + NEW_B39)) < 1e-6


def test_canh_bao_khi_xuat_ban_da_ghi(run):
    assert "Enable Editing" in run["export_warning"] and "số không mất" in run["export_warning"]


def test_nut_pdf_tai_dung_file(run):
    assert run["pdf_name"] == "GIA-QC-01_da_dien_xuat.pdf" and run["pdf_error"] is False
    assert (HERE / "out" / "xuat_ban2.pdf").read_bytes()[:5] == b"%PDF-"


def test_so_lon_cot_hep_khong_sai_so(run):
    big, small, frac, neg = run["narrow"]
    assert "E+" in big and big.startswith("1.23"), big
    assert "E-" in small, small
    assert frac.startswith("3.1") and "#" not in frac, frac
    assert neg.startswith("-") and "#" not in neg, neg


def test_ban_cu_chi_doc(run):
    assert run["old_state"] == "Đang xem bản 1, chỉ đọc"
    assert run["old_b8"] != 500
    assert run["old_readonly"]["before"] == run["old_readonly"]["after"]
    assert run["old_has_save"] is False
    assert "Người Giả" in run["versions_text"] and "Sửa 1 ô" in run["versions_text"]


def test_khoi_phuc(run):
    assert "Các bản hiện có không bị xóa" in run["restore_dialog"]
    assert run["v3_state"] == "Bản 3" and run["v3_b8"] != 500 and "nội dung bản 1" in run["restored_notice"]


def test_tai_lai_ban_xuat_chua_qua_excel(run):
    assert run["reupload_mismatch"] is False
    assert abs(run["reupload_b39"] - run["v2_b39"]) < 1e-9
    assert "." in run["reupload_display"]


def test_o_khoa_khong_sua_duoc_va_co_ly_do(run):
    assert run["locked"]["before"] == run["locked"]["after"]
    assert "khóa" in run["locked"]["notice"] and "không bị đổi" in run["locked"]["notice"]
