"""Gọi dịch vụ llm (Q135). api không có key; llm chỉ trả chữ thô, số token, hoặc một mã lỗi.

Không ghi payload, câu trả lời hay mã lỗi kèm nội dung vào log. Kết quả từ llm được coi là dữ liệu không đáng tin.
"""
from __future__ import annotations

import json
import os
import urllib.error
import urllib.request

LLM_URL = os.environ.get("LLM_URL", "http://llm:8100")
# Phải lớn hơn hạn chờ OpenAI của llm (45 giây) cộng thời gian llm xử lý, để api không cắt trước llm. Xem backend/tests/test_timeouts.py.
TIMEOUT_SECONDS = float(os.environ.get("LLM_TIMEOUT_SECONDS", "60"))
MAX_RESPONSE_BYTES = 2 * 1024 * 1024
_CODE_CHARS = set("ABCDEFGHIJKLMNOPQRSTUVWXYZ_")


class LlmFailed(Exception):
    def __init__(self, code: str, not_sent: bool = False):
        super().__init__(code)
        self.code = code
        # llm báo chắc chắn chưa gửi byte nào tới nhà cung cấp (chưa mở được kết nối qua egress): không tính tiền (PM 14/11).
        self.not_sent = not_sent


def _count(v: object) -> int | None:
    return v if isinstance(v, int) and not isinstance(v, bool) and v >= 0 else None


def call(system: str, prompt: str, response_schema: dict, schema_name: str, max_output_tokens: int) -> dict:
    """Trả {text, model, request_id, input_tokens, output_tokens, reasoning_tokens, latency_ms}. Lỗi: LlmFailed(mã)."""
    data = json.dumps({"system": system, "prompt": prompt, "response_schema": response_schema, "schema_name": schema_name,
                       "max_output_tokens": max_output_tokens}).encode()
    request = urllib.request.Request(f"{LLM_URL}/call", data=data, method="POST", headers={"Content-Type": "application/json"})
    try:
        with urllib.request.build_opener(urllib.request.ProxyHandler({})).open(request, timeout=TIMEOUT_SECONDS) as resp:
            raw = resp.read(MAX_RESPONSE_BYTES + 1)
    except urllib.error.HTTPError as exc:
        exc.close()
        raise LlmFailed("CHAT_BAD_REQUEST" if exc.code == 400 else "CHAT_MODEL_UNAVAILABLE") from None
    except TimeoutError:
        raise LlmFailed("CHAT_MODEL_TIMEOUT") from None
    except (urllib.error.URLError, OSError) as exc:
        raise LlmFailed("CHAT_MODEL_TIMEOUT" if isinstance(getattr(exc, "reason", None), TimeoutError) else "CHAT_MODEL_UNAVAILABLE") from None
    if len(raw) > MAX_RESPONSE_BYTES:
        raise LlmFailed("CHAT_MODEL_FAILED")
    try:
        body = json.loads(raw)
    except ValueError:
        raise LlmFailed("CHAT_MODEL_FAILED") from None
    if not isinstance(body, dict):
        raise LlmFailed("CHAT_MODEL_FAILED")
    if body.get("ok") is not True:
        code = body.get("code")
        # Chỉ nhận mã dạng CHAT_..., không chuyển tiếp chữ khác.
        ok_code = isinstance(code, str) and code.startswith("CHAT_") and len(code) <= 40 and set(code) <= _CODE_CHARS
        raise LlmFailed(code if ok_code else "CHAT_MODEL_FAILED", not_sent=body.get("not_sent") is True)
    if not isinstance(body.get("text"), str):
        raise LlmFailed("CHAT_MODEL_FAILED")
    return {
        "text": body["text"],
        "model": body.get("model") if isinstance(body.get("model"), str) else None,
        "request_id": body.get("request_id") if isinstance(body.get("request_id"), str) else None,
        "input_tokens": _count(body.get("input_tokens")),
        "output_tokens": _count(body.get("output_tokens")),
        "reasoning_tokens": _count(body.get("reasoning_tokens")),
        "latency_ms": _count(body.get("latency_ms")),
        "prompt_sha256": body.get("prompt_sha256") if isinstance(body.get("prompt_sha256"), str) and len(body["prompt_sha256"]) == 64 else None,
    }
