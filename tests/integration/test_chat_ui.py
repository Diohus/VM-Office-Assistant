"""Lát 2 (Q135): giao diện khung chat. Xem browser/chat.mjs, docs/product/acceptance_lat2.md. Model là máy giả fakeai."""
from __future__ import annotations

import json
import os
import pathlib
import subprocess

import pytest

from conftest import BASE, PROJECT, ROOT

HERE = pathlib.Path(__file__).parent
OUT = HERE / "out"
TEXTS = json.loads((ROOT / "vmai" / "src" / "vmai" / "data" / "doc_codes.json").read_text(encoding="utf8"))["chat_texts"]
NEVER_SEND = json.loads((ROOT / "fixtures" / "eval" / "chat_cases.json").read_text(encoding="utf8"))["never_send"]
AI_MARK = "#e9dcf7"


@pytest.fixture(scope="module")
def run():
    subprocess.run(["docker", "exec", f"{PROJECT}-db-1", "psql", "-U", "vm", "-d", "vm", "-c", "DELETE FROM chat_call"], capture_output=True, timeout=60)
    node = "node.exe" if os.name == "nt" else "node"
    proc = subprocess.run([node, "chat.mjs"], cwd=HERE / "browser", capture_output=True, text=True, encoding="utf8", timeout=580, env={**os.environ, "VM_BASE": BASE})
    assert proc.returncode == 0, proc.stderr[-2500:]
    return json.loads((OUT / "chat.json").read_text(encoding="utf8"))


def test_khong_loi_trang_khong_ra_ngoai(run):
    assert run["page_errors"] == [] and run["outbound"] == []


def test_chu_thich_mau_va_khong_phai_ban_thu(run):
    assert run["legend"].strip() == TEXTS["CHAT_UI_LEGEND"]
    # N5: chú thích chỉ hiện khi đang có ô tím: chưa có lượt nào, và sau khi Lưu (hết tím), thì không có.
    assert run["legend_before"] is False and run["legend_after_save"] is False
    assert run["trial"] is False


def test_2_1_2_10_2_12_xem_truoc_liet_ke_sheet_khong_sheet_an_khong_bieu_do(run):
    assert run["first"] == "preview"
    names = run["preview"]["sheets"]
    assert run["preview"]["sheet_lines"][0] == TEXTS["CHAT_UI_PREVIEW_SHEET"].replace("{sheet}", "KQ_QC_Thang").replace("{danh sách cột}", "A1:I47")
    assert names[0] == "KQ_QC_Thang" and names[1] == "GioiHan"
    assert "Huong_dan" not in names and "Bieu_do" not in names
    assert "Huong_dan" not in run["preview"]["prompt"]
    assert len(run["preview"]["hash"]) == 12


def test_2_4_2_11_khong_gui_ten(run):
    text = run["preview"]["prompt"]
    for never in NEVER_SEND:
        assert never not in text
    assert run["preview"]["not_sent"] and run["preview"]["not_sent"].startswith(TEXTS["CHAT_UI_PREVIEW_NOT_SENT"].split("{")[0])


def test_2_3_khong_gui_thi_khong_goi(run):
    assert run["not_sent"]["sends"] == 0 and run["not_sent"]["turns"] == 0
    assert run["not_sent"]["b10"] == run["edit"]["b10_before"]


def test_3_1_3_2_3_7_ap_co_dau_tim_tom_tat_khong_cau_model(run):
    e = run["edit"]
    assert run["again"] == "preview", "chưa đồng ý gửi lần nào thì vẫn hiện xem trước"
    assert e["b10"] == 5.62 and e["bg"].lower() == AI_MARK and e["ai"] == ["KQ_QC_Thang!B10"]
    assert e["h10"] != 7
    assert e["turn"]["edited"] == "1" and e["turn"]["summary"] == ["B10"]
    assert e["turn"]["skipped"] and e["turn"]["skipped"][0].startswith("H10:")
    assert e["turn"]["model"] is None, "Q135 (e): lượt có sửa ô không hiện câu của model"
    assert "[câu tạm]" not in e["turn"]["text"]


def test_3_4_hoan_tac_ca_luot_lam_lai_dua_lai(run):
    assert run["undo"]["b10"] == run["edit"]["b10_before"] and run["undo"]["ai"] == []
    assert (run["undo"]["bg"] or "").lower() != AI_MARK
    assert run["redo"]["b10"] == 5.62 and run["redo"]["bg"].lower() == AI_MARK and run["redo"]["ai"] == ["KQ_QC_Thang!B10"]


def test_2_5_3_8_luot_hai_khong_xem_truoc_cau_model_khi_hoi_lai(run):
    assert run["second"] == "turn"
    a = run["ask"]
    assert TEXTS["CHAT_ASK"] in a["text"] and a["model"] and a["model"].startswith(TEXTS["CHAT_UI_MODEL_SAID"])


def test_q135_e2_bi_chan_het_thi_khong_hien_cau_model(run):
    b = run["blocked"]
    assert run["blocked_kind"] == "turn"
    assert b["h10"] == b["h10_before"]
    assert TEXTS["CHAT_UI_EDITED_NONE"] in b["turn"]["text"]
    assert b["turn"]["skipped"] and b["turn"]["skipped"][0].startswith("H10:")
    assert b["turn"]["model"] is None and "Đã ghi xong." not in b["turn"]["text"], "model nói đã ghi trong khi bảng không đổi"


def test_2_8_cau_co_ten_khong_goi_model(run):
    p = run["person"]
    assert run["person_kind"] == "turn" and p["sends"] == 0
    assert p["turn"]["codes"] == ["CHAT_MESSAGE_HAS_PERSON"]
    assert TEXTS["CHAT_MESSAGE_HAS_PERSON"] in p["turn"]["text"]


def test_1_2_1_3_cho_va_huy(run):
    assert run["waiting"] == {"input_disabled": True, "send_disabled": True}
    # N12: lúc đã gửi cho AI thì câu chờ là CHAT_UI_WAITING.
    assert run["waiting_text"] == TEXTS["CHAT_UI_WAITING"]
    c = run["cancel"]
    assert TEXTS["CHAT_UI_STOPPED"] in c["turn"]["text"] and c["wait"] is False
    assert c["ai_after"] == c["ai_before"], "kết quả về muộn không được áp"


def test_doc_duoc_ngan_hoan_tac_cua_univer(run):
    """turnState đọc _undoStacks/_redoStacks (protected) của Univer. Nâng bản Univer mà hỏng ca này: sửa turnState, xem run.md."""
    assert run["undo_stacks_readable"] is True, "Univer đổi trường _undoStacks/_redoStacks: nút Hoàn tác lượt đang đoán theo dấu tím"


def test_hoan_tac_luot_an_nut_va_bao_da_hoan_tac(run):
    """PM 13/11: lượt đã hoàn tác thì ẩn nút, hiện CHAT_UI_TURN_UNDONE; Làm lại thì nút bấm được lại."""
    u = run["undo"]["turn"]
    assert u["button"] is None and u["undone"] == TEXTS["CHAT_UI_TURN_UNDONE"]
    assert TEXTS["CHAT_UI_UNDO_TURN_DISABLED"] not in u["text"], "lượt đã hoàn tác không phải là có sửa khác chen vào"
    r = run["redo"]["turn"]
    assert r["button"] == "enabled" and r["undone"] is None


def test_co_sua_khac_sau_luot_thi_nut_khoa_trong_khoa_that(run):
    """Gõ tay sau lượt AI: nút khóa, trông bị khóa (mờ, con trỏ không phải bàn tay), có câu CHAT_UI_UNDO_TURN_DISABLED."""
    t = run["typed"]["turn"]
    assert t["button"] == "disabled" and t["undone"] is None
    assert t["opacity"] < 1 and t["cursor"] != "pointer"
    assert TEXTS["CHAT_UI_UNDO_TURN_DISABLED"] in t["text"]


def test_3_5_go_tay_o_ai_thi_mat_dau(run):
    t = run["typed"]
    assert t["b10"] == 5.7 and (t["bg"] or "").lower() != AI_MARK and t["ai"] == []
    assert run["typed_undo"]["b10"] == 5.62 and run["typed_undo"]["ai"] == ["KQ_QC_Thang!B10"], "Hoàn tác lần gõ tay trả lại cả dấu"


def test_q135_h2_khong_ghi_de_o_co_chu(run):
    c = run["col_i"]
    assert c["i8"] == "chạy lại" and c["i9"] == "chạy lại"
    assert c["i10"] == run["note"], "I10 có chữ chưa gửi cho AI: phải giữ nguyên"
    assert c["i11"] == run["short_note"], "Q135 (h2): ghi chú ngắn đã gửi cho AI cũng phải giữ nguyên"
    for cell in ("I10", "I11"):
        codes = {s.split(":", 1)[1] for s in c["turn"]["skipped"] if s.startswith(f"{cell}:")}
        assert codes, f"{cell} phải có câu báo"
    assert not set(c["ai"]) & {"KQ_QC_Thang!I10", "KQ_QC_Thang!I11"}
    assert {"KQ_QC_Thang!I8", "KQ_QC_Thang!I9"} <= set(c["ai"])
    assert "[câu tạm]" not in c["turn"]["text"]


def test_5_1_5_2_5_4_luu_hoi_xac_nhan_lich_su_ghi_so_o(run):
    assert TEXTS["CHAT_UI_SAVE_CONFIRM"].replace("{N}", "3") in run["save_confirm"]
    assert run["after_review"] == "Bản 1", "Xem lại thì không lưu"
    s = run["saved"]
    assert s["state"] == "Bản 2" and s["b10"] == 5.62
    assert (s["bg"] or "").lower() != AI_MARK and s["ai"] == [], "5.4: mở lại bản đã lưu thì không còn dấu"
    assert "3 ô do AI sửa qua chat" in run["versions"]
    # K3 (5.4): lịch sử nêu ô nào do AI sửa, không chỉ số ô.
    assert TEXTS["CHAT_UI_HISTORY"].replace("{N}", "3").replace("{danh sách ô}", "B10, I8, I9") in run["versions"]
