"""Dịch vụ llm, không mạng: đọc key, dựng payload, đọc phản hồi, lỗi chỉ là mã."""
from __future__ import annotations

import io
import json
import urllib.error

import pytest

import llm_service as s

KEY = "unit-canary-key-0123456789"
REQ = {"system": "Hệ thống", "prompt": "Dữ liệu", "schema_name": "chat_edit", "response_schema": {"type": "object"}, "max_output_tokens": 6000}


@pytest.fixture
def secret(tmp_path, monkeypatch):
    f = tmp_path / "llm_env"
    f.write_text(f"ANTHROPIC_API_KEY=khac\n# ghi chu\nOPENAI_API_KEY={KEY}\n", encoding="utf-8")
    monkeypatch.setattr(s, "SECRET_FILE", f)
    return f


def test_doc_dung_dong_key(secret):
    assert s.read_key(secret) == KEY


@pytest.mark.parametrize("text", ["", "OPENAI_API_KEY=\n", "OTHER=1\n"])
def test_khong_co_key_la_ma(tmp_path, text):
    f = tmp_path / "e"
    f.write_text(text, encoding="utf-8")
    with pytest.raises(s.CallFailed) as err:
        s.read_key(f)
    assert err.value.code == "CHAT_MODEL_NO_KEY"


def test_file_secret_khong_co_la_ma(tmp_path):
    with pytest.raises(s.CallFailed) as err:
        s.read_key(tmp_path / "khong_co")
    assert err.value.code == "CHAT_MODEL_NO_KEY"


def test_payload_co_dinh_model_khong_luu_khong_tool():
    p = s.build_payload(REQ)
    assert p["model"] == "gpt-6-luna" and p["store"] is False and "tools" not in p
    assert p["text"]["format"]["strict"] is True and p["instructions"] == "Hệ thống"


@pytest.mark.parametrize("bad", [None, [], {**REQ, "max_output_tokens": 0}, {**REQ, "max_output_tokens": 9000}, {**REQ, "max_output_tokens": True},
                                 {**REQ, "prompt": ""}, {**REQ, "response_schema": "x"}, {k: v for k, v in REQ.items() if k != "system"}])
def test_yeu_cau_sai_bi_tu_choi(bad):
    with pytest.raises(s.CallFailed) as err:
        s.validate(bad)
    assert err.value.code == "CHAT_BAD_REQUEST"


def test_doc_phan_hoi_hop_le():
    out = s.parse_response({"id": "r1", "status": "completed", "model": "gpt-6-luna",
                            "output": [{"type": "reasoning"}, {"type": "message", "content": [{"type": "output_text", "text": "{\"a\":"}, {"type": "output_text", "text": "1}"}]}],
                            "usage": {"input_tokens": 10, "output_tokens": 5, "output_tokens_details": {"reasoning_tokens": 2}}})
    assert out == {"ok": True, "text": "{\"a\":1}", "model": "gpt-6-luna", "request_id": "r1", "input_tokens": 10, "output_tokens": 5, "reasoning_tokens": 2}


@pytest.mark.parametrize("body,code", [({"status": "incomplete"}, "CHAT_MODEL_TRUNCATED"), ({"status": "failed"}, "CHAT_MODEL_FAILED"),
                                       ({"status": "completed", "output": [{"type": "message", "content": [{"type": "refusal", "refusal": "x"}]}]}, "CHAT_MODEL_REFUSED")])
def test_phan_hoi_xau_thanh_ma(body, code):
    with pytest.raises(s.CallFailed) as err:
        s.parse_response(body)
    assert err.value.code == code


class _Opener:
    def __init__(self, result):
        self.result = result
        self.seen = None

    def open(self, request, timeout):
        self.seen = dict(request.header_items())
        if isinstance(self.result, Exception):
            raise self.result
        return io.BytesIO(self.result)


@pytest.mark.parametrize("status,code", [(401, "CHAT_MODEL_KEY_REJECTED"), (429, "CHAT_MODEL_BUSY"), (500, "CHAT_MODEL_UNAVAILABLE")])
def test_loi_http_khong_mang_than_loi(secret, status, code):
    err = urllib.error.HTTPError("https://x", status, "x", {}, io.BytesIO(f"Incorrect API key {KEY}".encode()))
    with pytest.raises(s.CallFailed) as caught:
        s.call(REQ, _Opener(err))
    assert caught.value.code == code and KEY not in repr(caught.value) and caught.value.__cause__ is None


def test_than_khong_phai_json(secret):
    with pytest.raises(s.CallFailed) as caught:
        s.call(REQ, _Opener(f"<html>{KEY}</html>".encode()))
    assert caught.value.code == "CHAT_MODEL_FAILED"


def test_goi_dat_key_vao_header_va_ket_qua_khong_co_key(secret):
    opener = _Opener(json.dumps({"status": "completed", "output": [], "usage": {}}).encode())
    out = s.call(REQ, opener)
    assert opener.seen["Authorization"] == f"Bearer {KEY}"
    assert KEY not in json.dumps(out)
