"""Không mất sửa, không xuất sai số (Q86, Q92), hoàn tác 20 bước, báo lệch số, dán từ Excel. Xem browser/guards.mjs."""
from __future__ import annotations

import io
import json
import os
import pathlib
import re
import subprocess
import zipfile

import openpyxl
import pdfplumber
import pytest

from conftest import BASE, FIX

HERE = pathlib.Path(__file__).parent
OUT = HERE / "out"
SHEET = "KQ_QC_Thang"


def make_mismatch_file() -> None:
    """Sửa số đã lưu của B39 trong bản sao, công thức giữ nguyên: giống file bị sửa ngoài Excel."""
    src = FIX / "GIA-QC-01_da_dien.xlsx"
    wb = openpyxl.load_workbook(src)
    index = wb.sheetnames.index(SHEET) + 1
    part = f"xl/worksheets/sheet{index}.xml"
    out = io.BytesIO()
    with zipfile.ZipFile(src) as zi, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for name in zi.namelist():
            data = zi.read(name)
            if name == part:
                text = data.decode("utf8")
                text, n = re.subn(r'(<c r="B39"[^>]*>\s*<f[^>]*>[^<]*</f>\s*<v>)[^<]*(</v>)', r"\g<1>12345\g<2>", text)
                assert n == 1, "phải thấy đúng ô B39"
                data = text.encode("utf8")
            zo.writestr(name, data)
    OUT.mkdir(exist_ok=True)
    (OUT / "lech_so.xlsx").write_bytes(out.getvalue())


@pytest.fixture(scope="module")
def run():
    make_mismatch_file()
    node = "node.exe" if os.name == "nt" else "node"
    proc = subprocess.run([node, "guards.mjs"], cwd=HERE / "browser", capture_output=True, text=True, encoding="utf8", timeout=580, env={**os.environ, "VM_BASE": BASE})
    assert proc.returncode == 0, proc.stderr[-2500:]
    return json.loads((OUT / "guards.json").read_text(encoding="utf8"))


def test_khong_loi_trang(run):
    assert run["page_errors"] == []


def test_o_dang_go_do_tinh_la_sua_chua_luu(run):
    assert run["midtype_editing"] is True
    assert run["midtype_kept"] == 7.5, "ở lại thì số đang gõ còn nguyên"
    assert "da_dien" in run["midtype_name"], "vẫn ở file cũ"


def test_o_dang_go_so_nhap_nhang_thi_dung_lai_va_bao(run):
    amb = run["midtype_ambiguous"]
    assert amb["dialog"] is False and amb["editing"] is True and "1.234" in amb["notice"]


def test_q92_xuat_luu_truoc_dung_so_dang_thay(run):
    assert run["q92_saved_state"] == "Bản 2"
    assert "Enable Editing" in run["q92_saved_warning"], "bản vừa lưu phải có cảnh báo, không ghi là giống file đã tải lên"
    values = openpyxl.load_workbook(OUT / "q92_luu.xlsx")[SHEET]
    assert values["B8"].value == 500 and values["B20"].value == 7.5
    with pdfplumber.open(OUT / "q92_luu.pdf") as pdf:
        text = "\n".join(p.extract_text() or "" for p in pdf.pages)
    shown = f"{run['q92_grid_b39']:.2f}"  # dấu chấm như lưới (Q122)
    assert shown in text, f"PDF phải có B39 đang thấy trên lưới: {shown}"


def test_q92_bo_thay_doi_ve_ban_da_luu_khong_trang(run):
    assert run["q92_drop_b9"] == run["b9_original"] and run["q92_drop_b8"] == 500 and run["q92_drop_dirty"] == 0
    with pdfplumber.open(OUT / "q92_bo.pdf") as pdf:
        assert pdf.pages


def test_q92_xem_ban_cu_khong_mat_sua(run):
    assert run["q92_stay_b10"] == 700
    assert "chỉ đọc" in run["q92_old_state"] and run["q92_old_b8"] == run["b8_original"]
    assert run["q92_back_state"] == "Bản 3" and run["q92_back_b10"] == 700


def test_q92_khoi_phuc_khi_co_sua(run):
    assert run["q92_restore_error"] is False
    assert run["q92_restore_state"] == "Bản 5", "lưu thành bản 4, khôi phục thành bản 5"
    assert run["q92_restore_b8"] == run["b8_original"]


def test_loi_may_chu_khong_bao_la_mat_ket_noi(run):
    # Q124: phân biệt lỗi mạng với lỗi máy chủ trả về.
    s400, s502 = run["server_400"], run["server_502"]
    assert "máy chủ của app trả lỗi (mã 400)" in s400["error"] and "mất kết nối" not in s400["error"] and s400["dirty"] > 0
    assert "mất kết nối" in s502["error"] and s502["dirty"] > 0


def test_qua_5000_o_bao_ngay_khong_gui_di(run):
    many = run["too_many"]
    assert many["posts"] == 0, "không gửi lên máy chủ"
    assert "quá 5000 ô sửa" in many["error"] and "chưa ghi" in many["error"] and "vẫn còn trên màn hình" in many["error"]
    assert "mất kết nối" not in many["error"]
    assert many["dirty"] == 7500 and many["state"] == "Bản 1"


def test_luu_loi_khong_mo_file_kia_va_con_sua(run):
    fail = run["save_fail"]
    assert fail["same_doc"] is True and fail["b13"] == 123 and fail["dirty"] > 0 and fail["error"] == "Giả lập lưu lỗi. Các sửa của anh/chị vẫn còn trên màn hình, chưa được lưu."


def test_hoan_tac_het_roi_mo_file_khac(run):
    assert run["undo_all_dirty"] == 0 and run["undo_all_dialog"] is False and "mau" in run["undo_all_name"]


def test_hoan_tac_lam_lai_20_buoc_dung_tung_buoc(run):
    undo = run["undo20"]
    assert undo["wrong_undo"] == 0 and undo["wrong_redo"] == 0 and undo["end_b39"] == undo["expected_end_b39"]


def test_lech_so_excel_bao_dung_o(run):
    bad = run["mismatch"]
    assert [m["cell"] for m in bad["list"]] == ["B39"] and bad["list"][0]["saved"] == 12345
    assert bad["text"], "phải có thông báo trên màn hình"


def test_dan_tu_excel_dung_tung_o_khong_lech_nghin_lan(run):
    paste = run["paste_excel"]
    assert paste["values"] == [1234567, 1234, 1.234, 5.47, 5.47, 1234.5, -1234, 0.5, "abc", 1234]
    # 7 ô có mẫu định dạng riêng: chỉ có số đang hiện. 2 ô General (5.47, 1234) được chữ thường của clipboard chứng minh (Q113).
    assert "7 ô có thể đã bị làm tròn" in paste["notice"] and "có thể hiểu hai cách" not in paste["notice"], "có mẫu định dạng nên không phải đoán dấu, nhưng vẫn báo có thể làm tròn"


def test_dan_chu_thuan_suy_va_bao_mot_lan(run):
    a = run["paste_text_evidence"]
    assert a["values"] == [1234567, 1234, 5.47] and "1 ô" in a["notice"] and "dấu chấm là số lẻ" in a["notice"]
    b = run["paste_text_none"]
    # Không có gì để suy: theo dấu của app, dấu chấm (Q122).
    assert b["values"] == [1234, 1.234] and "2 ô" in b["notice"] and "không có gì để biết chắc" in b["notice"] and "dấu chấm là số lẻ" in b["notice"]


def test_dan_chu_khong_thanh_cong_thuc(run):
    assert not run["paste_formula"]["formula"], "chữ dán từ ngoài không thành công thức"


def test_dan_vao_o_khoa_bi_tu_choi(run):
    locked = run["paste_locked"]
    assert locked["after"] == locked["before"] and locked["notice"] and "khóa" in locked["notice"]


def test_ctrl_s_la_luu_va_khong_mo_hop_trinh_duyet(run):
    assert run["ctrl_s"]["state"] == "Bản 2" and run["ctrl_s"]["prevented"] >= 1 and run["ctrl_s"]["b8"] == 6.5
    assert run["ctrl_s_midtype"]["state"] == "Bản 3" and run["ctrl_s_midtype"]["b10"] == 6.7, "ô đang gõ dở cũng được lưu"
    assert run["ctrl_s_nothing"]["state"] == "Bản 3" and run["ctrl_s_nothing"]["notice"], "không có sửa thì không tạo bản"


def test_thanh_tren_bao_co_sua_chua_luu(run):
    assert run["bar_clean"] is False and run["bar_dirty"] is True and run["bar_after_save"] is False
    assert run["bar_dirty_again"]["shown"] is True and run["bar_dirty_again"]["saved_notice"] is None, "không còn ghi Đã lưu bản N"
    assert run["bar_after_undo"] is False


def test_thong_bao_sau_khi_luu_dung_nguyen_van(run):
    assert run["ctrl_s"]["saved"] == "Đã lưu bản 2 trong app, ghi 1 ô. File trong folder của anh/chị chưa đổi. Muốn có file mới, bấm Xuất."


def test_hop_xuat_co_cau_pdf_va_bieu_do(run):
    first, second = run["export_text"]
    assert first == "Người nhận chỉ cần xem: bấm Tải PDF."
    assert "biểu đồ vẽ từ các ô đó hiện 0" in second and "Enable Editing" in second and second.startswith("Tải file Excel:")


def test_gia_lap_lech_pdf_bao_o_va_khong_xuat(run):
    assert "chưa xuất PDF" in run["pdf_fake_mismatch"]


def test_quay_lai_tu_man_chao_dung_ban_moi_nhat(run):
    assert "Bản 3" in run["back_row"] and "GIA-QC-01_da_dien" in run["back_row"]
    assert run["back"]["state"] == "Bản 3" and run["back"]["b8"] == 6.5 and run["back"]["b10"] == 6.7


def test_tai_lai_file_goc_la_tai_lieu_moi_ban_cu_con(run):
    assert run["reupload_original"]["state"] == "Bản 1" and run["reupload_original"]["b8"] == run["b8_original"]
    assert run["reupload_kept"] is True


def test_f2_roi_enter_khong_go_gi_khong_doi_so(run):
    f2 = run["f2"]
    assert f2["after"] == f2["before"] == 3.14159265358979 and f2["dirty"] == 0
    assert f2["ambiguous_after"] == f2["ambiguous_before"] == 1.234 and f2["ambiguous_notice"] is None and f2["ambiguous_dirty"] == 0


def test_f2_co_go_thi_van_la_sua(run):
    assert run["f2"]["typed_after"] == 5.5512
    amb = run["f2"]["typed_ambiguous"]
    assert amb["value"] == 5.58 and amb["editing"] is True and "5.581" in amb["notice"], "5.58 thêm số 1 là 5.581: có thể hiểu hai cách, phải hỏi, số cũ còn nguyên"


def test_f2_so_dai_go_them_thi_hoi_truoc_khi_cat_chu_so(run):
    # Q122: ô đang sửa chỉ hiện 12 chữ số. Không được âm thầm cắt các chữ số chưa hiện.
    h = run["hidden"]
    first = h["first"]
    assert first["value"] == 3.14159265358979 and first["editing"] is True
    assert "3.14159265358979" in first["notice"] and "3.14159265359" in first["notice"] and "3.141592653591" in first["notice"] and "chưa ghi" in first["notice"]
    changed = h["changed"]
    assert changed["value"] == 3.14159265358979 and changed["editing"] is True and "3.1415926535912" in changed["notice"], "gõ thêm sau khi hỏi thì hỏi lại"
    second = h["second"]
    assert second["value"] == 3.1415926535912 and second["editing"] is False and second["notice"] is None
    assert h["direct"]["value"] == 9 and h["direct"]["notice"] is None, "gõ thẳng thì không hỏi"


def test_chu_so_lan_lon_khong_de_thu_vien_doan(run):
    unclear = run["unclear"]
    assert unclear["value"] == 5.37 and unclear["editing"] is True and "không phải một số rõ ràng" in unclear["notice"]


def test_thong_bao_nhap_nhang_tu_dong_khi_da_ghi(run):
    close = run["ambiguous_close"]
    assert close["first"] and close["after_fix"] is None and close["value"] == 1234


def test_dan_so_hien_lech_so_that_phai_bao_cho_ca_vung(run):
    paste = run["paste_rounded"]
    # Giá trị thật trong Excel: 1234.5678, 0.1234, 5.4712, 1234.5678, 5.47 (canh giữa), 5.47 (canh trái), 1234. Trình duyệt chỉ nhận được số đang hiện.
    assert paste["values"] == [1235, 0.12, 5.47, 1234.57, 5.47, 5.47, 1234, "01/09/2026"]
    assert "4 ô có thể đã bị làm tròn" in paste["notice"] and "không phải số gốc" in paste["notice"] and "General" in paste["notice"]
    assert "2 ô được canh riêng" in paste["notice"] and "1 ô app không đọc được" in paste["notice"]
    rounded = [p for p in paste["places"] if "có thể đã làm tròn" in p]
    assert [p.split(":")[0].split("ô ")[1] for p in rounded] == ["L8", "L9", "L10", "L11"], "từng ô bị làm tròn phải có tên"
    assert "đã dán 1235" in rounded[0] and "đã dán 0.12" in rounded[1]


def test_dan_o_general_cot_hep_lay_so_goc_tu_chu_thuong(run):
    # Excel thật, cột hẹp: A hiện 1235, 0, -2.7, 5.47 (canh giữa), -12, 1234, 50%; B hiện 1234567890, 0.25.
    paste = run["paste_general"]
    assert paste["values"] == [[1234.9, 1234567890], [0.000123456, 0.25], [-2.71828, None], [5.4712, None], [-12.345, None], [1234, None], [0.5, None]]
    rounded = sorted(p.split(":")[0].split("ô ")[1] for p in paste["places"] if "có thể đã làm tròn" in p)
    # O8: 10 ký tự, Excel có thể đã cắt. N9: 11 ký tự. N14: mẫu 0%.
    assert rounded == ["N14", "N9", "O8"]
    assert "3 ô có thể đã bị làm tròn" in paste["notice"] and "1 ô được canh riêng" in paste["notice"]


def test_dan_excel_mat_chu_thuong_thi_bao_moi_o_so(run):
    paste = run["paste_general_no_plain"]
    rounded = [p for p in paste["places"] if "có thể đã làm tròn" in p]
    assert len(rounded) == 9 and "9 ô có thể đã bị làm tròn" in paste["notice"]


def test_dan_chu_thuan_dau_phay_va_chu_so_lan_lon_co_bao(run):
    # Q125, dòng 3.11: dán từ Notepad báo giống gõ tay.
    p = run["paste_plain_swapped"]
    assert p["values"] == [5.5, 0.5, 12.75, "1,5,5"]
    assert "3 ô viết số lẻ bằng dấu phẩy" in p["notice"] and '"5,5" là 5.5' in p["notice"]
    assert "1 ô không phải một số rõ ràng" in p["notice"] and '"1,5,5"' in p["notice"]
    assert any('ô K32: "12,75" là 12.75' in x for x in p["places"]) and any('ô K33: "1,5,5" không phải số rõ ràng' in x for x in p["places"])


def test_bam_dong_thong_bao_roi_go_ngay_khong_mat_phim(run):
    assert run["close_then_type"] == 12


def test_dan_o_dau_thang_thi_khong_dan_gi(run):
    narrow = run["paste_narrow"]
    assert narrow["value"] is None and "dấu thăng" in narrow["notice"]


def test_pdf_kho_02_da_sua_ra_duoc_va_dung_so(run):
    assert run["kho_pdf_error"] is None
    with pdfplumber.open(OUT / "kho_xuat.pdf") as pdf:
        text = "\n".join(p.extract_text() or "" for p in pdf.pages)
    # Chữ của C50 đè lên ô B50 trong PDF nên đọc ra xen kẽ, kiểm phần đuôi.
    assert "234.50" in text and "234,50" not in text, "TEXT phải giống lưới: 1,234.50"
    assert "#VALUE!" not in text, "VALUE không được lỗi"
    value_line = next(line for line in text.splitlines() if line.startswith("VALUE "))
    assert value_line.endswith("12.5"), "VALUE ra 12.5, số dùng dấu chấm như lưới (Q122)"


def test_hop_xuat_file_doi_cu_khong_noi_giong_het(run):
    assert run["legacy_export"][0].startswith("Bản này chưa sửa trong app. File tải về là bản .xlsx")
    assert "giống hệt" not in " ".join(run["legacy_export"])
