"""OpenRouter, trần chi phí, đo chi phí (Q55). Không gọi mạng ra ngoài; SDK thật chạy với máy chủ giả trên 127.0.0.1."""
from __future__ import annotations

import json
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer
from types import SimpleNamespace

import pytest

from vmai.contract import LlmCall
from vmai.costs import cost_of_form, summarize_by_type
from vmai.llm.base import (
    Capability,
    CapabilityError,
    CapabilityStatus,
    ConfigError,
    LLMBadOutput,
    LLMRequest,
    LLMResponse,
    LLMUnavailable,
    Message,
    MissingAPIKey,
    Profile,
    TextPart,
)
from vmai.llm.budget import Budget, BudgetedProvider, BudgetExceeded, Price, estimate_input_tokens
from vmai.llm.factory import create_provider, load_profile
from vmai.llm.fake import FakeProvider
from vmai.llm.openrouter_adapter import DEFAULT_ROUTING, OpenRouterProvider, build_payload, check_model_id

MODEL = "hang-gia/model-gia-1"
CAPS = {c: CapabilityStatus.YES for c in Capability}
REQUEST = LLMRequest(purpose="t", system="s", messages=(Message("user", (TextPart("x"),)),), max_output_tokens=100, response_schema={"type": "object"})


def _profile(**kw):
    return Profile("or", "openrouter", kw.pop("model", MODEL), "TEST_KEY", kw.pop("base_url", None), capabilities=CAPS, **kw)


# ---------- Không tự chọn model ----------
@pytest.mark.parametrize("bad", ["", "openrouter/auto", "auto", "a/b:auto", "a/b,c/d"])
def test_model_must_be_one_explicit_id(bad):
    with pytest.raises(ConfigError):
        check_model_id(bad)


def test_openrouter_profile_has_no_model_until_pm_chooses():
    with pytest.raises(ConfigError):
        load_profile("openrouter")


def test_payload_and_routing_defaults_turn_off_substitution():
    assert DEFAULT_ROUTING == {"allow_fallbacks": False, "data_collection": "deny", "require_parameters": True}
    payload = build_payload(_profile(), REQUEST)
    assert payload["model"] == MODEL and "models" not in payload and "route" not in payload
    assert payload["response_format"]["type"] == "json_schema"
    assert payload["messages"][0]["role"] == "system"


def _raw(model=MODEL, content='{"a":1}', finish="stop", cost=None):
    usage = SimpleNamespace(prompt_tokens=3, completion_tokens=4, cost=cost)
    message = SimpleNamespace(content=content, tool_calls=None, refusal=None)
    return SimpleNamespace(model=model, id="r1", usage=usage, choices=[SimpleNamespace(message=message, finish_reason=finish)])


def _client(raw, capture=None):
    def create(**kw):
        if capture is not None:
            capture.update(kw)
        return raw

    return SimpleNamespace(chat=SimpleNamespace(completions=SimpleNamespace(create=create)))


def test_response_from_another_model_is_rejected():
    with pytest.raises(LLMBadOutput):
        OpenRouterProvider(_profile(), client=_client(_raw(model="hang-khac/model-khac"))).generate(REQUEST)


def test_routing_is_sent_and_reported_cost_is_kept():
    capture = {}
    out = OpenRouterProvider(_profile(), client=_client(_raw(cost=0.0012), capture)).generate(REQUEST)
    assert capture["extra_body"]["provider"]["allow_fallbacks"] is False
    assert out.cost_usd == 0.0012 and out.parsed() == {"a": 1}


def test_sdk_error_does_not_repeat_content():
    def boom(**kw):
        raise RuntimeError("noi dung nhay cam")

    client = SimpleNamespace(chat=SimpleNamespace(completions=SimpleNamespace(create=boom)))
    with pytest.raises(LLMUnavailable) as exc:
        OpenRouterProvider(_profile(), client=client).generate(REQUEST)
    assert "nhay cam" not in str(exc.value)


def test_missing_key_stops_before_any_call(monkeypatch):
    monkeypatch.delenv("TEST_KEY", raising=False)
    with pytest.raises(MissingAPIKey):
        OpenRouterProvider(_profile()).generate(REQUEST)


def test_real_sdk_round_trip_against_local_server(monkeypatch):
    received = []

    class Handler(BaseHTTPRequestHandler):
        def do_POST(self):
            body = self.rfile.read(int(self.headers["content-length"]))
            received.append({"path": self.path, "body": json.loads(body)})
            data = json.dumps(
                {
                    "id": "gen-1",
                    "object": "chat.completion",
                    "created": 1,
                    "model": MODEL,
                    "choices": [{"index": 0, "finish_reason": "stop", "message": {"role": "assistant", "content": '{"a":1}'}}],
                    "usage": {"prompt_tokens": 3, "completion_tokens": 4, "total_tokens": 7, "cost": 0.0001},
                }
            ).encode()
            self.send_response(200)
            self.send_header("content-type", "application/json")
            self.send_header("content-length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)

        def log_message(self, *args):
            pass

    httpd = HTTPServer(("127.0.0.1", 0), Handler)
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    monkeypatch.setenv("TEST_KEY", "khong-phai-key-that")
    monkeypatch.setenv("NO_PROXY", "127.0.0.1")
    try:
        out = OpenRouterProvider(_profile(base_url=f"http://127.0.0.1:{httpd.server_port}")).generate(REQUEST)
    finally:
        httpd.shutdown()
    assert out.parsed() == {"a": 1} and out.input_tokens == 3
    assert received[0]["path"].endswith("/chat/completions")
    assert received[0]["body"]["provider"]["data_collection"] == "deny"
    assert received[0]["body"]["model"] == MODEL


# ---------- Trần chi phí ----------
def _fake(**usage):
    profile = Profile("f", "fake", "fake-1", capabilities=CAPS)
    return FakeProvider(profile, handler=lambda req: LLMResponse("fake", "fake-1", text="{}", **usage))


PRICE = Price(input_per_mtok=1.0, output_per_mtok=2.0, as_of="2026-11-04", source="giả")


def test_token_cap_blocks_before_calling():
    inner = _fake(input_tokens=1, output_tokens=1)
    provider = BudgetedProvider(inner, Budget(max_tokens_per_run=50))
    with pytest.raises(BudgetExceeded):
        provider.generate(REQUEST)
    assert inner.requests == []


def test_usd_cap_blocks_before_calling():
    inner = _fake(input_tokens=1, output_tokens=1)
    provider = BudgetedProvider(inner, Budget(max_tokens_per_run=10**6, max_usd_per_run=0.0000001, price=PRICE))
    with pytest.raises(BudgetExceeded):
        provider.generate(REQUEST)
    assert inner.requests == []


def test_usd_cap_without_price_refuses_unless_verifying():
    inner = _fake(input_tokens=1, output_tokens=1)
    with pytest.raises(BudgetExceeded):
        BudgetedProvider(inner, Budget(max_tokens_per_run=10**6, max_usd_per_run=1.0)).generate(REQUEST)
    BudgetedProvider(inner, Budget(max_tokens_per_run=10**6, max_usd_per_run=1.0, allow_missing_price=True)).generate(REQUEST)


def test_budget_accumulates_and_stops_a_run():
    provider = BudgetedProvider(_fake(input_tokens=100, output_tokens=100), Budget(max_tokens_per_run=450))
    provider.generate(REQUEST)
    provider.generate(REQUEST)
    assert provider.tokens_used == 400
    with pytest.raises(BudgetExceeded):
        provider.generate(REQUEST)


def test_missing_usage_counts_as_worst_case_not_zero():
    provider = BudgetedProvider(_fake(), Budget(max_tokens_per_run=10**6))
    provider.generate(REQUEST)
    assert provider.tokens_used >= estimate_input_tokens(REQUEST) + REQUEST.max_output_tokens


def test_capability_error_comes_before_budget():
    profile = Profile("f", "fake", "fake-1", capabilities={})
    with pytest.raises(CapabilityError):
        BudgetedProvider(FakeProvider(profile), Budget(max_tokens_per_run=1)).generate(REQUEST)


def test_real_profiles_cannot_be_created_without_a_budget(monkeypatch):
    profile = load_profile("claude")
    assert profile.budget is not None and profile.budget.max_tokens_per_run > 0
    provider = create_provider(profile)
    assert isinstance(provider, BudgetedProvider)
    with pytest.raises(ConfigError):
        create_provider(Profile("x", "openai", "m", "K"))
    # Giá chưa xác minh thì chưa gọi được, kể cả khi đã bật khả năng.
    assert profile.budget.price is None and profile.budget.max_usd_per_run is not None


# ---------- Đo chi phí ----------
def _call(**kw):
    return LlmCall(purpose="t", provider="p", model="m", **kw)


def test_no_llm_call_is_zero_and_labeled_so():
    cost = cost_of_form("f1", "excel", [])
    assert cost.usd == 0 and cost.basis == "khong_goi_model"


def test_cost_basis_is_honest():
    assert cost_of_form("f", "t", [_call(input_tokens=1, output_tokens=1, cost_usd=0.5)]).basis == "do"
    assert cost_of_form("f", "t", [_call(input_tokens=10**6, output_tokens=0)], PRICE).basis == "tinh_theo_bang_gia"
    missing = cost_of_form("f", "t", [_call(input_tokens=1, output_tokens=1)])
    assert missing.usd is None and missing.basis == "thieu_so_lieu"
    assert cost_of_form("f", "t", [_call()], PRICE).basis == "thieu_so_lieu"


def test_summary_does_not_mix_form_types():
    summary = summarize_by_type([cost_of_form("a", "excel", []), cost_of_form("b", "word", [_call(input_tokens=1, output_tokens=1, cost_usd=0.5)])])
    assert summary["excel"]["median_usd"] == 0 and summary["word"]["median_usd"] == 0.5
