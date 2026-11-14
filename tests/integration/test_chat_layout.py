"""Lát 2, Owner thử 13/11 (O1, O2): acceptance_lat2.md mục 1.5, 1.6, 1.7. Xem browser/chat_layout.mjs. Model là máy giả fakeai.

Cửa sổ 1366x657 (màn 1366x768 của Owner, trừ thanh trình duyệt), thu phóng 100%. Mọi cú bấm đi bằng chuột theo tọa độ.
"""
from __future__ import annotations

import json
import os
import pathlib
import re
import subprocess

import pytest

from conftest import BASE, PROJECT, ROOT

HERE = pathlib.Path(__file__).parent
OUT = HERE / "out"
TEXTS = json.loads((ROOT / "vmai" / "src" / "vmai" / "data" / "doc_codes.json").read_text(encoding="utf8"))["chat_texts"]


@pytest.fixture(scope="module")
def run():
    subprocess.run(["docker", "exec", f"{PROJECT}-db-1", "psql", "-U", "vm", "-d", "vm", "-c", "DELETE FROM chat_call"], capture_output=True, timeout=60)
    node = "node.exe" if os.name == "nt" else "node"
    proc = subprocess.run([node, "chat_layout.mjs"], cwd=HERE / "browser", capture_output=True, text=True, encoding="utf8", timeout=580, env={**os.environ, "VM_BASE": BASE})
    result = json.loads((OUT / "chat_layout.json").read_text(encoding="utf8"))
    assert proc.returncode == 0, (result.get("failed"), proc.stderr[-1500:])
    return result


def test_khong_loi_trang_khong_ra_ngoai(run):
    assert run["page_errors"] == [] and run["outbound"] == []


def test_1_5_moi_nut_bam_duoc_bang_chuot(run):
    """Nút nằm trọn trong cửa sổ và trong khung chat; phần tử ở tọa độ bấm đúng là nút đó, không phải ô gõ hay phần khác."""
    bad = [c for c in run["clicks"] if not (c["found"] and c["in_view"] and c["hit_ok"] and not c["disabled"] and c["in_chat"] in (True, None))]
    assert bad == [], bad
    labels = {c["label"] for c in run["clicks"]}
    for need in ("Đồng ý gửi", "Không gửi (chưa chọn gì)", "Hủy", "Hoàn tác lượt này", "Hoàn tác lượt này (30 ô)", "Đóng khung chat", "mở lại khung chat"):
        assert need in labels, f"chưa thử bấm: {need}"


def test_1_5_xem_truoc_dai_than_cuon_nut_van_trong_khung(run):
    assert run["long_kind"] == "preview"
    assert run["long"]["body_scrolls"], "ca thử phải có màn xem trước dài hơn khung"
    assert run["long"]["sheets"][:2] == ["KQ_QC_Thang", "GioiHan"]


def test_1_5_bam_dong_y_gui_thi_ap_va_hoan_tac_duoc(run):
    a = run["after_send"]
    assert a["b10"] == 5.62 and a["turn"]["summary"] == 1
    assert run["after_undo_b10"] == a["b10_before"], "bấm Hoàn tác lượt này bằng chuột phải bỏ cả lượt"


def test_1_5_cau_tra_loi_dai_va_tom_tat_30_o(run):
    assert run["long_ask_kind"] == "turn" and run["long_ask"]["model"]
    assert run["many"]["summary"] == 30
    assert run["many_undo_b8"] != 4.01, "Hoàn tác lượt 30 ô bằng chuột phải bỏ cả lượt"


def test_1_5_huy_dong_mo_khung(run):
    assert TEXTS["CHAT_UI_STOPPED"] in run["cancel"]["turn"]["text"] and run["cancel"]["wait"] is False
    assert run["closed"] and run["reopened"]


def test_1_6_dang_xem_truoc_khong_co_nut_khoa_gay_nham(run):
    long = run["long"]
    assert long["main_send"] is False and long["input"] is False, "đang xem trước thì không có nút Gửi chính hay ô gõ bị khóa"
    buttons = run["long_buttons"]
    assert all(not b["disabled"] for b in buttons), buttons
    assert [b["text"] for b in buttons] == ["Đóng", TEXTS["CHAT_UI_PREVIEW_SEND"], TEXTS["CHAT_UI_PREVIEW_CANCEL"]]
    assert TEXTS["CHAT_UI_PREVIEW_SEND"] == "Đồng ý gửi"
    # Nút bị khóa trong danh sách lượt (lượt cũ có sửa khác chen vào) phải trông bị khóa thật.
    for b in run["turn_disabled_look"]:
        assert b["opacity"] < 1 and b["cursor"] != "pointer", b


@pytest.mark.parametrize("label", ["chưa chọn gì", "bấm chuột một ô", "một ô B10", "ô cột H (H10)", "ô gộp tiêu đề A1", "cả cột H",
                                   "tiêu đề một cột (chuột)", "tiêu đề một hàng (chuột)", "cả sheet (góc trên trái)"])
def test_1_7_moi_cach_chon_deu_toi_xem_truoc(run, label):
    s = run["selections"][label]
    assert s["kind"] == "preview", s.get("turn")
    # Vùng ghi dạng A1:I47, không phải "B:B" hay "1:65" (nguyên nhân O2).
    assert re.fullmatch(r"Sheet KQ_QC_Thang: [A-Z]+\d+:[A-Z]+\d+\.", s["sheets"][0]), s["sheets"][0]
    assert s["preview_gone"] and s["input_back"].startswith("Glucose mức 1"), "Không gửi thì đóng xem trước, trả lại câu đã gõ"


def test_1_7_ca_sheet_la_vung_dang_dung(run):
    assert run["selections"]["cả sheet (góc trên trái)"]["sheets"][0] == "Sheet KQ_QC_Thang: A1:I47."
    assert run["selections"]["cả cột H"]["sheets"][0] == "Sheet KQ_QC_Thang: H1:H47."


def test_o2_cong_thuc_chua_tinh_khong_bi_tu_choi(run):
    """Ô công thức chưa có kết quả (bộ tính còn chạy): trước đây gửi v null, vmai trả CHAT_BAD_INPUT."""
    assert run["pending_formula_value"] is None, "ca thử phải có ô công thức chưa có kết quả"
    assert run["selections"]["công thức chưa tính"]["kind"] == "preview"


def test_n7_o_da_dung_san_khong_tinh_vao_khong_ap(run):
    a = run["already"]
    assert run["already_kind"] in ("turn", "preview")
    assert a["turn"]["summary"] == 1, "chỉ B10 được áp"
    assert a["turn"]["already"] == ["B12:CHAT_NO_CHANGE"]
    assert not any(s.startswith("B12:") for s in a["turn"]["skipped"]), "ô đã đúng không nằm trong danh sách không áp"
    partial = TEXTS["CHAT_UI_PARTIAL"].split("{")[0]
    assert partial not in a["turn"]["text"], "chỉ có ô đã đúng thì không có câu áp K, không áp M"
    assert a["b12_after"] == a["b12"]


def test_n4_bieu_mau_chua_dien_duoc_thi_khong_noi_app_mo_dien_tu_file(run):
    f = run["fill"]
    assert run["fill_kind"] == "turn"
    assert TEXTS["CHAT_UNSUPPORTED_FILL_FROM_FILE"] not in f["turn"]["text"], "app không mở Điền từ file trên biểu mẫu này"
    assert TEXTS["CHAT_UNSUPPORTED"] in f["turn"]["text"]
    assert f["fill_panel"] is False
    assert "chỉ điền được biểu mẫu QC mẫu" in f["page"], "băng của app nêu lý do"


def test_n2_chon_vung_so_gui_kem_tieu_de_va_cot_ngay_chi_de_doc(run):
    """ChatEditPlan 0.2: chọn B8:C37 thì gửi kèm A1:C37 chỉ để đọc; AI ghi ô ngoài vùng chọn (A10) thì bị chặn."""
    assert run["n2_kind"] == "preview"
    p = run["n2_preview"]
    assert p["sheets"][0] == "Sheet KQ_QC_Thang: B8:C37."
    assert p["context"] == [TEXTS["CHAT_UI_PREVIEW_CONTEXT"].replace("{vùng}", "A1:C37").replace("{vùng sửa}", "B8:C37")]
    assert "A10:" in p["prompt"], "cột Ngày của vùng chọn phải có trong nội dung gửi"
    t = run["n2"]["turn"]
    assert t["summary"] == 1 and run["n2"]["b10"] == 5.62
    assert "A10:CHAT_CELL_READ_ONLY" in t["skipped"]
    assert run["n2"]["a10"] == run["n2"]["a10_before"]


def test_k4_o_ngoai_than_bang_van_ap_nhung_tach_nhom_canh_bao(run):
    o = run["outside"]
    stats = [f"I{r}" for r in range(38, 45)]
    assert o["outside"] == stats, o
    assert o["normal"] == ["I8"], o
    # PO/BA 4.16: mỗi ô kèm nhãn dòng (row_label của vmai); dòng 44 trống thì chỉ có địa chỉ ô.
    labels = ["Số lần chạy", "Trung bình", "SD", "CV%", "Lệch so với đích %", "Số lần ngoài ±2SD"]
    listed = ", ".join([f"{c} ({l})" for c, l in zip(stats, labels)] + ["I44"])
    assert o["warn"] == TEXTS["CHAT_UI_OUTSIDE_BODY"].replace("{N}", "7").replace("{danh sách ô}", listed)
    assert o["i38_formula"] and "F38" in o["i38_formula"], "ô ngoài thân bảng vẫn được áp"


def test_n12_luc_chuan_bi_khong_noi_ai_dang_lam(run):
    p = run["preparing"]
    assert p["phase"] == "prepare" and p["text"] == TEXTS["CHAT_UI_PREPARING"]
    assert p["text"] != TEXTS["CHAT_UI_WAITING"]


def test_6_5_loi_khi_gui_tra_cau_ve_o_go(run):
    r = run["retry"]
    assert "CHAT_MODEL_UNAVAILABLE" in r["turn"]["codes"]
    assert r["input"] == r["message"], "câu vừa gửi phải trở lại ô gõ nguyên văn để bấm Gửi lại"


def test_6_5_loi_do_nguoi_dung_giu_cau_de_sua(run):
    p = run["person"]
    assert run["person_kind"] == "turn" and p["turn"]["codes"] == ["CHAT_MESSAGE_HAS_PERSON"]
    assert p["input"] == p["message"]


def test_2_13_sheet_phu_khong_vua_tran_thi_khong_gui(run):
    snap = run["big_snapshot"]
    assert "DuLieu_Nam" not in snap["sheets"] and "DuLieu_Nam" in snap["skipped"]
    assert snap["sheets"][:2] == ["KQ_QC_Thang", "GioiHan"] and snap["total"] <= 3000
    assert run["big_kind"] == "preview", run["big_preview"]
    assert "DuLieu_Nam" not in run["big_preview"]["sheets"]
