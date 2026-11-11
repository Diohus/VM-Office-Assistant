"""T-DEV-09 mốc 2: giao diện lát 1 nối AI (Q133). Xem browser/fill.mjs, docs/product/acceptance_lat1.md."""
from __future__ import annotations

import json
import os
import pathlib
import subprocess

import pytest

from conftest import BASE

HERE = pathlib.Path(__file__).parent
OUT = HERE / "out"


@pytest.fixture(scope="module")
def run():
    node = "node.exe" if os.name == "nt" else "node"
    proc = subprocess.run([node, "fill.mjs"], cwd=HERE / "browser", capture_output=True, text=True, encoding="utf8", timeout=580, env={**os.environ, "VM_BASE": BASE})
    assert proc.returncode == 0, proc.stderr[-2500:]
    return json.loads((OUT / "fill.json").read_text(encoding="utf8"))


def test_khong_loi_trang_khong_ra_ngoai(run):
    assert run["page_errors"] == [] and run["outbound"] == [], "1.3: 0 yêu cầu ra ngoài"


def test_chi_bieu_mau_qc_mau(run):
    assert "chỉ điền được biểu mẫu QC mẫu" in run["not_qc"]["error"] and run["not_qc"]["dialog"] is False


def test_1_1_1_2_1_5_hop_hoi_lai_va_huy(run):
    c = run["confirm"]
    assert "GIA-QC-01_mau.xlsx" in c and "GIA-NGUON_2026-09.xlsx" in c, "nêu cả hai tên file"
    assert "Bản thử này chỉ dùng với file giả đội gửi kèm. Anh/chị đừng chọn file kết quả thật của khoa." in c
    assert run["cancel"] == {"panel": False, "state": "Bản 1", "open_batches": []}


def test_6_3_co_sua_chua_luu_thi_hoi(run):
    assert run["unsaved_asks"] is True
    assert run["d6"]["apply_unsaved_asks"] is True and run["d6"]["state_after"] == "Bản 2", "bấm Ghi khi có sửa chưa lưu: hỏi, không ghi"


def test_3_1_3_3_dau_khung_duyet(run):
    x = run["x01"]
    assert x["header"].startswith("Bản thử này chưa đánh dấu ô nào là ✓ Khớp file đã chọn.") and "GIA-NGUON_2026-09.xlsx" in x["header"] and "{" not in x["header"]
    assert x["source"] == "File nguồn: GIA-NGUON_2026-09.xlsx"
    assert x["counts"] == "Còn chờ: 180 ô vàng, 0 ô đỏ."
    assert x["mapping"] == 6 and x["rows"] == 30


def test_4_1_de_xuat_chua_ghi_khong_vao_luoi(run):
    assert run["x01"]["b8_before"] is None and run["x01"]["b8_still"] is None


def test_3_10_bo_nhan_lam_giam_so_o(run):
    x = run["x01"]
    assert x["apply_label"] == "Ghi 0 ô thành bản mới"
    assert x["after_5"] == "Ghi 30 ô thành bản mới" and x["after_unaccept"] == "Ghi 24 ô thành bản mới"


def test_3_4_3_5_3_5a_phim(run):
    k = run["keys"]
    assert k["grid_enter"] == "Ghi 30 ô thành bản mới", "Enter trong lưới không nhận dòng nào"
    assert k["modifiers"] == "Ghi 30 ô thành bản mới", "Ctrl+A, Ctrl+Enter, Shift+Enter không nhận"
    assert k["hold_enter"] == "Ghi 36 ô thành bản mới", "giữ Enter chỉ nhận một dòng"


def test_4_2_4_3_ghi_thanh_ban_moi(run):
    x = run["x01"]
    a = x["after_apply"]
    assert a["state"] == "Bản 2" and a["b8"] == 5.51
    assert a["notice"] == "Đã ghi 36 ô từ file GIA-NGUON_2026-09.xlsx thành bản 2, còn 144 ô chưa ghi."
    assert x["versions"][0]["summary"] == "Nhận 36 ô từ file GIA-NGUON_2026-09.xlsx" and x["versions"][0]["author"] == "Người Giả"


def test_5_3_d10_bao_ly_do(run):
    assert run["d10"] and "đang chờ duyệt trong một lần điền khác" in run["d10"]


def test_3_8_mo_lai_lan_dien_con_cho(run):
    r = run["reopen"]
    assert "Còn một lần điền từ file GIA-NGUON_2026-09.xlsx chưa xong" in r["notice"]
    assert r["counts"] == "Còn chờ: 144 ô vàng, 0 ô đỏ."


def test_d4_x04_chon_lan_doc(run):
    x = run["x04"]
    assert "5.64" in x["red_text"] and "5.94" in x["red_text"] and "Chọn lần đọc 2: 5.94" in x["red_text"]
    assert x["label"] == "Ghi 1 ô thành bản mới"
    assert x["after"]["value"] == 5.94
    assert x["after"]["versions"][0]["summary"] == "Nhận 1 ô từ file X04_mot_ngay_hai_lan_chay.xlsx; ô đỏ: 1 chọn lần đọc 2"


def test_d4_x10_tu_go_nhap_nhang_bi_bao_va_de_trong(run):
    x = run["x10"]
    assert x["reds"] == 5
    assert "có thể hiểu hai cách" in x["ambiguous"] and x["label_after_ambiguous"] == "Ghi 0 ô thành bản mới"
    assert x["label"] == "Ghi 1 ô thành bản mới"
    assert x["after"]["versions"][0]["summary"] == "Nhận 1 ô từ file X10_so_vo_ly.xlsx; ô đỏ: 1 tự gõ, 1 để trống"


def test_d6_5_1_5_2_o_ghi_de_nam_rieng(run):
    d = run["d6"]
    assert "Ô đã có số" in d["section"] and "(B8)" in d["section"] and "9.99 → 5.51" in d["section"]
    assert d["label_after_row"] == "Ghi 5 ô thành bản mới" and d["b8_choice"] is None, "nhận dòng không nhận ô ghi đè"


def test_6_2_mat_ket_noi_khi_tao_lo(run):
    n = run["net"]
    assert "mất kết nối" in n["error"] and n["panel"] is False and n["state"] == "Bản 1"


def test_mo_tai_lieu_khac_thi_khung_cu_dong(run):
    # PO/BA 11/11 (lỗi chặn lát 1): trước đây khung của A còn lại khi mở B, Ghi ghi vào A.
    o = run["other_doc"]
    assert o["panel_before"] is True and o["same_doc"] is False
    assert o["panel_after"] is False and o["apply_button"] is False
    assert o["versions_a"] == 1 and o["versions_b"] == 1
    assert o["open_a"] == 1, "lô của A vẫn chờ, mở lại A thì có câu nhắc"


def test_n8_lo_0_o_noi_ro_khong_co_o_nao(run):
    n = run["n8"]
    assert n["empty"] and "Không có ô nào điền được" in n["empty"] and n["header"] is False and n["apply"] is False
    assert "10/2026" in n["notices"] and "{" not in n["notices"]
    assert n["open"] == 0, "không còn câu Còn một lần điền"


def test_k2_ngay_thieu_viet_kieu_viet_nam(run):
    t = run["k2"]["notices"]
    assert "/09" in t and "2026-09" not in t and "{" not in t, t


def test_k3_cot_nguoi_benh_cau_lat1(run):
    t = run["k3"]["notices"]
    assert "thông tin người bệnh" in t and "gửi ra ngoài" not in t and "{" not in t, t


def test_n5_bo_lan_dien_hoi_lai_va_khong_tao_ban_moi(run):
    n = run["n5"]
    assert "ô còn chờ sẽ bị bỏ" in n["ask"] and "đã ghi" in n["ask"] and "vẫn giữ" in n["ask"]
    assert n["after_cancel"] == {"panel": True, "open": 1}, "Hủy: lô còn nguyên"
    d = n["after_drop"]
    assert d["panel"] is False and d["versions"] == 1 and d["open"] == 0 and "Đã bỏ lần điền" in d["notice"]
    assert n["reopen_notice"] is None, "không còn câu Còn một lần điền"
    assert n["refill"] == {"panel": True, "error": None}, "lô mới chạm cùng ô mở được ngay"


def test_anh_o_do_co_o_do(run):
    assert run["x04"]["red_in_view"] is True
