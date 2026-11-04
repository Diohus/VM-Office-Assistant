"""Chạy SDK thật của hai adapter với máy chủ giả trên 127.0.0.1. Không có gì rời khỏi máy.

Chỉ chứng minh SDK chấp nhận tham số và adapter đọc được phản hồi theo đúng dạng tài liệu.
Không chứng minh model thật chạy được hay cho kết quả tốt.
"""
from __future__ import annotations

import json
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer

import pytest


from vmai.llm.anthropic_adapter import AnthropicProvider
from vmai.llm.base import ImagePart, LLMRefusal, LLMRequest, Message, Profile, TextPart, ToolSpec
from vmai.llm.openai_adapter import OpenAIProvider

SCHEMA = {"type": "object", "properties": {"a": {"type": "integer"}}, "required": ["a"], "additionalProperties": False}
REQUEST = LLMRequest(
    purpose="test",
    system="s",
    messages=(Message("user", (ImagePart(b"\x89PNG"), TextPart("x"))),),
    tools=(ToolSpec("t1", "d", SCHEMA),),
    response_schema=SCHEMA,
)
DUMMY_KEY = "khong-phai-key-that"


@pytest.fixture()
def server():
    received: list[dict] = []
    replies: list[dict] = []

    class Handler(BaseHTTPRequestHandler):
        def do_POST(self):
            body = self.rfile.read(int(self.headers["content-length"]))
            received.append({"path": self.path, "body": json.loads(body), "headers": {k.lower(): v for k, v in self.headers.items()}})
            data = json.dumps(replies.pop(0)).encode()
            self.send_response(200)
            self.send_header("content-type", "application/json")
            self.send_header("content-length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)

        def log_message(self, *args):
            pass

    httpd = HTTPServer(("127.0.0.1", 0), Handler)
    thread = threading.Thread(target=httpd.serve_forever, daemon=True)
    thread.start()
    yield f"http://127.0.0.1:{httpd.server_port}", received, replies
    httpd.shutdown()


def _caps():
    from vmai.llm.base import Capability, CapabilityStatus

    return {c: CapabilityStatus.YES for c in Capability}


def test_openai_sdk_round_trip(server, monkeypatch):
    url, received, replies = server
    monkeypatch.setenv("TEST_KEY", DUMMY_KEY)
    replies.append(
        {
            "id": "resp_1",
            "object": "response",
            "created_at": 1,
            "model": "gpt-6-luna",
            "status": "completed",
            "output": [
                {"type": "message", "id": "m1", "role": "assistant", "status": "completed", "content": [{"type": "output_text", "text": '{"a":1}', "annotations": []}]},
                {"type": "function_call", "id": "f1", "call_id": "c1", "name": "t1", "arguments": '{"a":2}', "status": "completed"},
            ],
            "usage": {"input_tokens": 3, "output_tokens": 4, "total_tokens": 7, "input_tokens_details": {"cached_tokens": 0}, "output_tokens_details": {"reasoning_tokens": 0}},
            "parallel_tool_calls": True,
            "tool_choice": "auto",
            "tools": [],
        }
    )
    provider = OpenAIProvider(Profile("p", "openai", "gpt-6-luna", "TEST_KEY", url, capabilities=_caps()))
    out = provider.generate(REQUEST)
    assert out.parsed() == {"a": 1}
    assert out.tool_calls[0].arguments == {"a": 2}
    assert out.input_tokens == 3
    assert received[0]["path"].endswith("/responses")
    assert received[0]["body"]["store"] is False
    assert received[0]["headers"]["authorization"] == f"Bearer {DUMMY_KEY}"


def test_anthropic_sdk_round_trip_and_refusal(server, monkeypatch):
    url, received, replies = server
    monkeypatch.setenv("TEST_KEY", DUMMY_KEY)
    # Đặt các biến này thì SDK có thể tự dùng; adapter phải dùng đúng key trong cấu hình.
    monkeypatch.setenv("ANTHROPIC_API_KEY", "key-khac-khong-duoc-dung")
    replies.append(
        {
            "id": "msg_1",
            "type": "message",
            "role": "assistant",
            "model": "claude-opus-5-5",
            "stop_reason": "tool_use",
            "stop_sequence": None,
            "content": [{"type": "text", "text": '{"a":1}'}, {"type": "tool_use", "id": "c1", "name": "t1", "input": {"a": 2}}],
            "usage": {"input_tokens": 3, "output_tokens": 4},
        }
    )
    provider = AnthropicProvider(Profile("p", "anthropic", "claude-opus-5-5", "TEST_KEY", url, capabilities=_caps()))
    out = provider.generate(REQUEST)
    assert out.parsed() == {"a": 1}
    assert out.tool_calls[0].arguments == {"a": 2}
    assert received[0]["path"].endswith("/v1/messages")
    # SDK 2.x gửi key qua Authorization. Quan trọng là đúng key trong cấu hình, không phải key khác trên máy.
    assert DUMMY_KEY in received[0]["headers"]["authorization"]
    assert "fallbacks" not in received[0]["body"]
    assert received[0]["body"]["output_config"]["format"]["type"] == "json_schema"

    replies.append(
        {
            "id": "msg_2",
            "type": "message",
            "role": "assistant",
            "model": "claude-opus-5-5",
            "stop_reason": "refusal",
            "stop_sequence": None,
            "stop_details": {"type": "refusal", "category": None, "explanation": None},
            "content": [],
            "usage": {"input_tokens": 3, "output_tokens": 0},
        }
    )
    with pytest.raises(LLMRefusal):
        provider.generate(REQUEST)
