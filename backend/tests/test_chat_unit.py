"""Khung chat lát 2 (Q135), phần không cần Docker: ngày giờ Việt Nam, trần chi phí, bản xem trước, một lượt mỗi tài liệu."""
from __future__ import annotations

import datetime as dt
import uuid

import pytest

from app import chat

UTC = dt.timezone.utc


def at(text: str) -> dt.datetime:
    return dt.datetime.fromisoformat(text)


@pytest.mark.parametrize("moment,day", [
    ("2026-11-12T16:59:59+00:00", "2026-11-12"),  # 23:59:59 giờ Việt Nam
    ("2026-11-12T17:00:00+00:00", "2026-11-13"),  # 00:00 giờ Việt Nam, UTC vẫn là ngày 12
    ("2026-11-12T23:30:00+00:00", "2026-11-13"),  # 06:30 giờ Việt Nam ngày 13
    ("2026-11-13T00:30:00+00:00", "2026-11-13"),  # đã qua 00:00 UTC nhưng vẫn ngày 13 giờ Việt Nam (7.7)
    ("2026-11-13T06:30:00+07:00", "2026-11-13"),
    ("2026-11-14T00:05:00+07:00", "2026-11-14"),
])
def test_ngay_theo_gio_viet_nam(moment, day):
    assert chat.vn_day(at(moment)).isoformat() == day


def settings(turn=0.02, day=1.0, now=None):
    return chat.Settings(turn, day, now)


@pytest.mark.parametrize("estimated,spent,reserved,code", [
    (None, 0, 0, "CHAT_NO_PRICE"),
    (0.021, 0, 0, "CHAT_TURN_LIMIT"),
    (0.01, 0.995, 0, "CHAT_DAY_LIMIT"),
    (0.01, 0.98, 0.015, "CHAT_DAY_LIMIT"),
])
def test_tran_chan_truoc_khi_goi(estimated, spent, reserved, code):
    with pytest.raises(chat.BudgetRefused) as err:
        chat.check_budget(estimated, spent, reserved, settings())
    assert err.value.code == code


def test_trong_tran_thi_cho_goi():
    chat.check_budget(0.02, 0.98, 0.0, settings())
    chat.check_budget(0.0033, 0.0, 0.0, settings())


def test_doc_bien_moi_truong(monkeypatch):
    for name in ("CHAT_MAX_USD_PER_TURN", "CHAT_MAX_USD_PER_DAY", "CHAT_TEST_NOW"):
        monkeypatch.delenv(name, raising=False)
    s = chat.load_settings()
    assert (s.max_usd_per_turn, s.max_usd_per_day, s.test_now, s.is_test) == (0.02, 1.0, None, False)
    monkeypatch.setenv("CHAT_MAX_USD_PER_DAY", "0.001")
    assert chat.load_settings().max_usd_per_day == 0.001 and chat.load_settings().is_test
    monkeypatch.setenv("CHAT_MAX_USD_PER_DAY", "")
    monkeypatch.setenv("CHAT_TEST_NOW", "2026-11-13T06:30:00+07:00")
    s = chat.load_settings()
    assert s.is_test and s.now() == at("2026-11-12T23:30:00+00:00") and chat.vn_day(s.now()).isoformat() == "2026-11-13"


@pytest.mark.parametrize("name,value", [("CHAT_TEST_NOW", "2026-11-13T06:30:00"), ("CHAT_MAX_USD_PER_DAY", "-1"), ("CHAT_MAX_USD_PER_TURN", "abc")])
def test_bien_sai_thi_khong_chay(monkeypatch, name, value):
    monkeypatch.setenv(name, value)
    with pytest.raises(ValueError):
        chat.load_settings()


def test_loi_co_the_bi_tinh_tien_ghi_bang_uoc_tinh():
    assert chat.cost_of_failure("CHAT_MODEL_TIMEOUT", 0.005) == 0.005
    assert chat.cost_of_failure("CHAT_MODEL_NO_KEY", 0.005) == 0.0
    assert chat.cost_of_failure("CHAT_MODEL_KEY_REJECTED", 0.005) == 0.0


PREPARED = {"prompt": "Dữ liệu"}


def test_ban_xem_truoc_chi_gui_mot_lan_va_het_han():
    now = [1000.0]
    store = chat.Previews(clock=lambda: now[0])
    doc = uuid.uuid4()
    key, item = store.put(doc, 3, PREPARED)
    assert len(key) >= 20 and item.prompt_sha256 and len(store) == 1
    assert store.take(key) is item
    assert store.take(key) is None  # không gửi lại được cùng một bản
    key2, _ = store.put(doc, 3, PREPARED)
    now[0] += chat.PREVIEW_SECONDS + 1
    assert store.take(key2) is None and len(store) == 0


def test_ban_xem_truoc_co_tran_so_luong():
    store = chat.Previews()
    keys = [store.put(uuid.uuid4(), 1, PREPARED)[0] for _ in range(chat.MAX_PREVIEWS + 5)]
    assert len(store) == chat.MAX_PREVIEWS
    assert store.take(keys[0]) is None and store.take(keys[-1]) is not None


def test_moi_tai_lieu_mot_luot_va_giu_cho_tien():
    turns = chat.Turns()
    a, b = uuid.uuid4(), uuid.uuid4()
    turns.start(a, 0.4, 0.0, settings(turn=1.0))
    with pytest.raises(chat.BudgetRefused) as err:
        turns.start(a, 0.001, 0.0, settings())
    assert err.value.code == "CHAT_BUSY"
    # Lượt của a đang giữ 0,4: b không được làm tổng vượt trần ngày dù chi thật mới 0,5.
    with pytest.raises(chat.BudgetRefused) as err:
        turns.start(b, 0.2, 0.5, settings(turn=1.0))
    assert err.value.code == "CHAT_DAY_LIMIT"
    turns.finish(a)
    turns.start(b, 0.2, 0.5, settings(turn=1.0))
    turns.finish(b)
    turns.start(a, 0.001, 0.0, settings())


def test_peek_kiem_tran_khong_giu_cho_tien():
    """Bước xem trước: cùng mã như bước gửi, tính cả tiền lượt khác đang giữ, nhưng không giữ chỗ và không chặn CHAT_BUSY."""
    turns = chat.Turns()
    a, b = uuid.uuid4(), uuid.uuid4()
    with pytest.raises(chat.BudgetRefused) as err:
        turns.peek(0.021, 0.0, settings())
    assert err.value.code == "CHAT_TURN_LIMIT"
    with pytest.raises(chat.BudgetRefused) as err:
        turns.peek(None, 0.0, settings())
    assert err.value.code == "CHAT_NO_PRICE"
    turns.start(a, 0.4, 0.0, settings(turn=1.0))
    with pytest.raises(chat.BudgetRefused) as err:
        turns.peek(0.2, 0.5, settings(turn=1.0))
    assert err.value.code == "CHAT_DAY_LIMIT"
    turns.peek(0.001, 0.0, settings(turn=1.0))  # cùng tài liệu đang chờ: xem trước vẫn được, bước gửi mới báo CHAT_BUSY
    turns.peek(0.001, 0.0, settings(turn=1.0))
    turns.finish(a)
    turns.start(b, 0.2, 0.5, settings(turn=1.0))  # peek không giữ chỗ tiền


def test_tien_loi_goi_model_chua_gui_thi_bang_0():
    """PM 14/11: llm báo chưa gửi byte nào thì không tính tiền; lỗi có thể đã gửi thì ghi ước tính xấu nhất."""
    assert chat.cost_of_failure("CHAT_MODEL_UNAVAILABLE", 0.0034, not_sent=True) == 0.0
    assert chat.cost_of_failure("CHAT_MODEL_TIMEOUT", 0.0034, not_sent=True) == 0.0
    assert chat.cost_of_failure("CHAT_MODEL_UNAVAILABLE", 0.0034) == 0.0034
    assert chat.cost_of_failure("CHAT_MODEL_TIMEOUT", 0.0034, not_sent=False) == 0.0034
    assert chat.cost_of_failure("CHAT_MODEL_KEY_REJECTED", 0.0034) == 0.0


@pytest.mark.parametrize("flag,expected", [(True, True), (False, False), ("true", False), (None, False)])
def test_llm_client_doc_co_not_sent(monkeypatch, flag, expected):
    """api chỉ tin not_sent khi llm trả đúng true (bool); mã lỗi người dùng thấy không đổi."""
    import io
    import json

    from app import llm_client

    body = {"ok": False, "code": "CHAT_MODEL_UNAVAILABLE"}
    if flag is not None:
        body["not_sent"] = flag

    class _Opener:
        def open(self, request, timeout):
            return io.BytesIO(json.dumps(body).encode())

    monkeypatch.setattr(llm_client.urllib.request, "build_opener", lambda *a: _Opener())
    with pytest.raises(llm_client.LlmFailed) as err:
        llm_client.call("s", "p", {}, "chat_edit", 100)
    assert err.value.code == "CHAT_MODEL_UNAVAILABLE" and err.value.not_sent is expected
