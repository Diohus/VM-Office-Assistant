"""Lát 2, mốc API (Q135): xem trước rồi gửi, chat_call không có nội dung, trần chi phí, một lượt mỗi tài liệu, ai_cells khi Lưu.

Model là máy giả fakeai (compose.llm-test.yaml), chọn câu trả lời theo dấu FAKE:... trong câu gõ. Không gọi mạng thật.
Dữ liệu: fixtures/templates/chat_lat2/GIA-QC-01_chat.xlsx (giả). Yêu cầu dựng như trình duyệt bằng request_from_xlsx của vmai.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import json
import subprocess
import sys
import threading

import pytest

from conftest import BASE, PROJECT, ROOT, request, upload

sys.path.insert(0, str(ROOT / "vmai" / "src"))
sys.path.insert(0, str(ROOT / "vmai" / "eval"))
from run_chat_gates import request_from_xlsx  # noqa: E402

CHAT_FILE = ROOT / "fixtures" / "templates" / "chat_lat2" / "GIA-QC-01_chat.xlsx"
CASES = json.loads((ROOT / "fixtures" / "eval" / "chat_cases.json").read_text(encoding="utf8"))
NEVER_SEND = CASES["never_send"]
CANARY_FILE = ROOT / "tests" / "integration" / "out" / "llm_canary.env"
VN = dt.timezone(dt.timedelta(hours=7))


def post(path: str, payload: dict | None = None):
    body = json.dumps(payload).encode() if payload is not None else None
    status, _, raw = request("POST", path, body, headers={"Content-Type": "application/json"} if payload is not None else {})
    return status, raw, json.loads(raw) if raw[:1] in (b"{", b"[") else None


def sql(query: str) -> str:
    out = subprocess.run(["docker", "exec", f"{PROJECT}-db-1", "psql", "-U", "vm", "-d", "vm", "-At", "-c", query],
                         capture_output=True, text=True, encoding="utf8", timeout=60)
    assert out.returncode == 0, out.stderr
    return out.stdout.strip()


@pytest.fixture(autouse=True)
def clean_calls():
    # Bộ test chạy nhiều lần mỗi ngày: không để tiền của lần trước làm vượt trần ngày.
    sql("DELETE FROM chat_call")
    yield
    sql("DELETE FROM chat_call")


@pytest.fixture
def doc():
    status, body = upload(CHAT_FILE.read_bytes(), CHAT_FILE.name)
    assert status == 201, body
    return body


def chat_request(doc: dict, message: str) -> dict:
    req = request_from_xlsx(CHAT_FILE, CASES["active_sheet"], CASES["region"], CASES["extra_sheets"], message)
    req["base_version"] = doc["version"]
    return req


def preview(doc: dict, message: str):
    return post(f"/api/documents/{doc['document_id']}/chat/preview", chat_request(doc, message))


def turn(doc: dict, message: str):
    status, raw, body = preview(doc, message)
    assert status == 200, raw[:300]
    return (*post(f"/api/chat/{body['preview_id']}/send"), body)


def test_xem_truoc_dung_chuoi_gui_di_va_khong_co_ten(doc):
    status, raw, body = preview(doc, "Glucose mức 1 ngày 3/9 là 5,62")
    assert status == 200
    assert body["prompt_sha256"] == hashlib.sha256(body["prompt"].encode()).hexdigest()
    assert body["masked"] and body["cell_count"] > 0 and body["region_cells"] <= 3000
    assert 0 < body["estimated_cost_usd"] < 0.02 and body["chat_test_settings"] is False
    text = raw.decode("utf8")
    for never in NEVER_SEND:
        assert never not in text and json.dumps(never)[1:-1] not in text, "giá trị không được gửi nằm trong bản xem trước"


def test_gui_dung_ban_da_xem_va_chat_call_khong_co_noi_dung(doc):
    status, raw, body, shown = turn(doc, "Sửa giúp tôi")
    assert status == 200, raw[:300]
    assert body["plan"]["status"] == "ask" and body["prompt_sha256"] == shown["prompt_sha256"]
    seen = subprocess.run(["docker", "logs", f"{PROJECT}-llm-1"], capture_output=True, text=True, encoding="utf8", errors="replace").stderr
    assert f"prompt={shown['prompt_sha256'][:12]}" in seen
    row = sql("SELECT status || '|' || cells_sent || '|' || masked_groups || '|' || input_tokens || '|' || CASE WHEN cost_usd > 0 THEN 't' ELSE 'f' END FROM chat_call")
    assert row == f"ask|{shown['cell_count']}|{len(shown['masked'])}|1200|t"
    columns = sql("SELECT string_agg(column_name, ',' ORDER BY column_name) FROM information_schema.columns WHERE table_name = 'chat_call'")
    assert not {"message", "prompt", "raw_text", "text", "value"} & set(columns.split(","))


def test_mot_ban_xem_truoc_chi_gui_mot_lan(doc):
    status, _, body = preview(doc, "Sửa giúp tôi")
    assert post(f"/api/chat/{body['preview_id']}/send")[0] == 200
    status, _, again = post(f"/api/chat/{body['preview_id']}/send")
    assert status == 404 and again["code"] == "CHAT_PREVIEW_EXPIRED"
    assert post("/api/chat/khong-co-ban-nay/send")[0] == 404


def test_cong_chay_tren_ket_qua_model(doc):
    status, raw, body, _ = turn(doc, "Glucose mức 1 ngày 3/9 là 5,62 FAKE:edit")
    assert status == 200, raw[:300]
    plan = body["plan"]
    assert plan["status"] == "edit"
    assert [(op["cell"], op["value"]["normalized"]) for op in plan["ops"]] == [("B10", "5.62")]
    assert [r["cell"] for r in plan["rejected_ops"]] == ["H10"]  # 7 không có trong câu gõ
    assert plan["summary"] and plan["summary"][0]["cell"] == "B10"
    assert sql("SELECT ops || '|' || array_to_string(rejected_codes, ',') FROM chat_call") == f"1|{plan['rejected_ops'][0]['reason']}"


def test_cau_go_co_ten_nguoi_khong_tao_ban_xem_truoc(doc):
    status, raw, body = preview(doc, "Hôm nay Trần Giả Định chạy, Glucose mức 1 ngày 5/9 là 5,40")
    assert status == 422 and body["code"] == "CHAT_MESSAGE_HAS_PERSON"
    assert "Trần Giả Định" not in raw.decode("utf8")
    assert sql("SELECT count(*) FROM chat_call") == "0"


@pytest.mark.parametrize("message,code", [("", "CHAT_MESSAGE_EMPTY"), ("x" * 1001, "CHAT_MESSAGE_TOO_LONG")])
def test_cau_go_sai_bi_tu_choi(doc, message, code):
    status, _, body = preview(doc, message)
    assert status == 422 and body["code"] == code


def test_ban_khong_co_thi_404(doc):
    req = chat_request(doc, "Sửa giúp tôi")
    req["base_version"] = 99
    assert post(f"/api/documents/{doc['document_id']}/chat/preview", req)[0] == 404


def _today_vn() -> str:
    return dt.datetime.now(VN).date().isoformat()


def _spend(doc: dict, day: str, cost: float) -> None:
    sql(f"INSERT INTO chat_call (document_id, at, vn_day, status, cells_sent, masked_groups, estimated_cost_usd, cost_usd) "
        f"VALUES ('{doc['document_id']}', now(), '{day}', 'ask', 1, 0, {cost}, {cost})")


def test_vuot_tran_ngay_thi_bao_ngay_o_xem_truoc(doc):
    """PO/BA lượt 2 (ca 1a): báo trần ngay ở bước xem trước, trước khi người dùng đọc và bấm đồng ý."""
    _spend(doc, _today_vn(), 0.999)
    status, _, body = preview(doc, "Sửa giúp tôi")
    assert status == 429 and body["code"] == "CHAT_DAY_LIMIT" and "0 giờ ngày mai (giờ Việt Nam)" in body["message"]
    assert "preview_id" not in body and "prompt" not in body
    assert sql("SELECT count(*) FROM chat_call") == "1"  # không có lượt gọi mới


def test_tien_tieu_them_giua_xem_truoc_va_gui_thi_buoc_gui_van_chan(doc):
    """Giữa xem trước và gửi, lượt khác có thể tiêu tiền: bước gửi vẫn kiểm lại trần, không gọi model."""
    status, raw, body = preview(doc, "Sửa giúp tôi")
    assert status == 200, raw[:300]
    _spend(doc, _today_vn(), 0.999)
    status, _, sent = post(f"/api/chat/{body['preview_id']}/send")
    assert status == 429 and sent["code"] == "CHAT_DAY_LIMIT"
    assert sql("SELECT count(*) FROM chat_call") == "1"


def test_tien_ngay_hom_qua_khong_tinh_vao_hom_nay(doc):
    yesterday = (dt.datetime.now(VN).date() - dt.timedelta(days=1)).isoformat()
    _spend(doc, yesterday, 5.0)
    assert turn(doc, "Sửa giúp tôi")[0] == 200


def test_tien_tinh_theo_bang_gia_tu_so_token(doc):
    status, _, body, _ = turn(doc, "Sửa giúp tôi FAKE:costly")
    assert status == 200
    # 50.000 token vào x 0,10 + 6.000 token ra x 0,50 (USD mỗi triệu token, vmai/src/vmai/config/llm.yaml)
    assert body["cost_usd"] == pytest.approx(0.008)
    assert float(sql("SELECT cost_usd FROM chat_call")) == pytest.approx(0.008)


@pytest.mark.parametrize("marker,code,charged", [("FAKE:echo_401", "CHAT_MODEL_KEY_REJECTED", False), ("FAKE:slow", "CHAT_MODEL_TIMEOUT", True)])
def test_loi_model_chi_la_ma_va_ghi_tien_than_trong(doc, marker, code, charged):
    status, raw, body, shown = turn(doc, f"Sửa giúp tôi {marker}")
    assert status == 502 and body["code"] == code and "Bảng không đổi" in body["message"]
    key = CANARY_FILE.read_text(encoding="utf8").strip().split("=", 1)[1]
    assert key not in raw.decode("utf8")
    cost = float(sql("SELECT cost_usd FROM chat_call WHERE status = 'model_error'"))
    assert cost == (pytest.approx(shown["estimated_cost_usd"]) if charged else 0.0)


def test_mot_tai_lieu_mot_luot_dang_cho(doc):
    _, _, slow = preview(doc, "Sửa giúp tôi FAKE:slow")
    _, _, second = preview(doc, "Sửa giúp tôi")
    result = {}
    worker = threading.Thread(target=lambda: result.update(first=post(f"/api/chat/{slow['preview_id']}/send")))
    worker.start()
    threading.Event().wait(1.5)  # lượt chậm đã vào tới llm (máy giả chờ 40 giây, llm hết hạn sau 5 giây)
    status, _, body = post(f"/api/chat/{second['preview_id']}/send")
    worker.join(30)
    assert status == 409 and body["code"] == "CHAT_BUSY"
    assert result["first"][0] == 502


def test_luu_ghi_ro_o_do_ai_sua(doc):
    did = doc["document_id"]
    base = {"base_version": doc["version"], "author": "Người thử", "edits": [{"sheet": "KQ_QC_Thang", "cell": "B10", "value": 5.62}]}
    status, _, body = post(f"/api/documents/{did}/versions", {**base, "ai_cells": [{"sheet": "KQ_QC_Thang", "cell": "C11"}]})
    assert status == 422 and body["code"] == "AI_CELLS_INVALID"
    status, raw, body = post(f"/api/documents/{did}/versions", {**base, "ai_cells": [{"sheet": "KQ_QC_Thang", "cell": "B10"}]})
    assert status == 201, raw[:300]
    assert "1 ô do AI sửa qua chat" in body["summary"]
    _, _, versions = request("GET", f"/api/documents/{did}/versions")
    rows = {v["version"]: v for v in json.loads(versions)}
    assert rows[body["version"]]["ai_cells"] == 1 and rows[doc["version"]]["ai_cells"] == 0
    # K3 (5.4): lịch sử nêu ô nào, chỉ địa chỉ ô.
    assert rows[body["version"]]["ai_cell_list"] == ["KQ_QC_Thang!B10"] and rows[doc["version"]]["ai_cell_list"] == []
    assert sql(f"SELECT by_ai FROM cell_edit WHERE document_id = '{did}' AND cell = 'B10'") == "t"


def test_health_bao_ban_that():
    _, _, raw = request("GET", "/api/health")
    assert json.loads(raw)["chat_test_settings"] is False
    assert BASE
