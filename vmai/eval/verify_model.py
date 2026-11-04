"""Xác minh một model theo docs/ai/gpt6_luna_checklist.md. Chỉ gọi mạng khi có --confirm-network.

Dùng --model không có: model phải nằm trong cấu hình, không tự đổi. Có trần token cho cả lần chạy.
"""
from __future__ import annotations

import argparse
import dataclasses
import json
import statistics
import time
from pathlib import Path

from vmai.llm.base import ImagePart, LLMError, LLMRequest, Message, TextPart, ToolSpec
from vmai.llm.budget import Budget
from vmai.llm.factory import create_provider, load_profile
from vmai.llm.keys import read_api_key

ROOT = Path(__file__).resolve().parents[2]
SAMPLE_IMAGE = ROOT / "fixtures" / "eval" / "generated" / "images" / "clean_s301" / "page.jpg"
OUT = ROOT / "vmai" / "out"
MAX_TOKENS = 40_000
_ANSWER_SCHEMA = {"type": "object", "properties": {"answer": {"type": "string"}}, "required": ["answer"], "additionalProperties": False}


def _text(prompt: str, **kw) -> LLMRequest:
    return LLMRequest(purpose="verify", system="Trả lời ngắn.", messages=(Message("user", (TextPart(prompt),)),), max_output_tokens=500, **kw)


def checks(provider, image: bytes | None) -> list[tuple[str, callable]]:
    def text():
        response = provider.generate(_text("Trả lời đúng một từ: xin chào"))
        return {"model": response.model, "text": response.text[:80], "input_tokens": response.input_tokens, "cost_usd": response.cost_usd}

    def structured():
        return {"parsed": provider.generate(_text("Điền answer là ok.", response_schema=_ANSWER_SCHEMA)).parsed()}

    def tool():
        spec = ToolSpec("ghi_nhan", "Ghi nhận một chuỗi", _ANSWER_SCHEMA)
        return {"tool_calls": [c.arguments for c in provider.generate(_text("Gọi tool ghi_nhan với answer là ok.", tools=(spec,))).tool_calls]}

    def vision():
        request = LLMRequest(
            purpose="verify",
            system="Trả lời ngắn.",
            messages=(Message("user", (ImagePart(image or b"", "image/jpeg"), TextPart("Dòng tiêu đề của ảnh ghi gì?"))),),
            response_schema=_ANSWER_SCHEMA,
            max_output_tokens=500,
        )
        response = provider.generate(request)
        return {"parsed": response.parsed(), "input_tokens": response.input_tokens, "output_tokens": response.output_tokens, "cost_usd": response.cost_usd}

    def latency():
        times = []
        for _ in range(5):
            started = time.perf_counter()
            provider.generate(_text("Trả lời: ok"))
            times.append(time.perf_counter() - started)
        return {"median_s": statistics.median(times), "max_s": max(times)}

    return [("text", text), ("structured_output", structured), ("tool_calls", tool), ("image_input", vision), ("latency", latency)]


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--profile", default=None)
    parser.add_argument("--confirm-network", action="store_true", help="Cho phép gọi dịch vụ ngoài, tốn phí, gửi ảnh giả")
    args = parser.parse_args()

    profile = load_profile(args.profile)
    print(f"Profile {profile.name}: {profile.provider}/{profile.model}, key từ {profile.api_key_env}, trần {MAX_TOKENS} token")
    print("Các bước:", ", ".join(name for name, _ in checks(None, None)))
    print(f"Ảnh giả: {SAMPLE_IMAGE}")
    if not args.confirm_network:
        print("Chưa có --confirm-network: không gọi mạng.")
        return

    read_api_key(profile)  # thiếu key thì dừng ở đây, trước khi gọi
    # Xác minh diễn ra trước khi biết giá, nên ở đây chỉ dùng trần token.
    provider = create_provider(dataclasses.replace(profile, budget=Budget(max_tokens_per_run=MAX_TOKENS)), allow_unverified=True)
    results = {}
    for name, run in checks(provider, SAMPLE_IMAGE.read_bytes()):
        try:
            results[name] = {"ok": True, **run()}
        except LLMError as exc:
            results[name] = {"ok": False, "error": f"{type(exc).__name__}: {exc}"}
        print(name, json.dumps(results[name], ensure_ascii=False))
    results["tokens_used"] = provider.tokens_used
    OUT.mkdir(exist_ok=True)
    (OUT / f"verify_{profile.name}.json").write_text(json.dumps(results, ensure_ascii=False, indent=1), encoding="utf-8")


if __name__ == "__main__":
    main()
