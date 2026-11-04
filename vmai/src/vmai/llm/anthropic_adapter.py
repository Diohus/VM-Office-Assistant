"""Adapter Claude, dùng Messages API qua SDK chính thức. Chưa từng gọi thật: chờ key."""
from __future__ import annotations

import base64
import time
from typing import Any

from .base import (
    ImagePart,
    LLMBadOutput,
    LLMProvider,
    LLMRefusal,
    LLMRequest,
    LLMResponse,
    LLMTruncated,
    LLMUnavailable,
    Message,
    Profile,
    TextPart,
    ToolCall,
)
from .keys import read_api_key


def _content(message: Message) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    for part in message.parts:
        if isinstance(part, TextPart):
            out.append({"type": "text", "text": part.text})
        elif isinstance(part, ImagePart):
            encoded = base64.b64encode(part.data).decode("ascii")
            out.append({"type": "image", "source": {"type": "base64", "media_type": part.media_type, "data": encoded}})
    return out


def build_payload(profile: Profile, request: LLMRequest) -> dict[str, Any]:
    # Cố ý không bật `fallbacks`: yêu cầu của dự án là không âm thầm đổi model.
    payload: dict[str, Any] = {
        "model": profile.model,
        "system": request.system,
        "messages": [{"role": m.role, "content": _content(m)} for m in request.messages],
        "max_tokens": request.max_output_tokens,
    }
    if request.tools:
        payload["tools"] = [
            {"name": t.name, "description": t.description, "input_schema": t.parameters, "strict": True}
            for t in request.tools
        ]
    if request.response_schema is not None:
        payload["output_config"] = {"format": {"type": "json_schema", "schema": request.response_schema}}
    return payload


class AnthropicProvider(LLMProvider):
    def __init__(self, profile: Profile, allow_unverified: bool = False, client: Any = None):
        super().__init__(profile, allow_unverified)
        self._client = client

    def _get_client(self) -> Any:
        if self._client is None:
            import anthropic

            # Truyền key rõ ràng: SDK không được tự tìm thông tin đăng nhập khác trên máy.
            self._client = anthropic.Anthropic(
                api_key=read_api_key(self.profile),
                base_url=self.profile.base_url,
                timeout=self.profile.timeout_seconds,
                max_retries=2,
            )
        return self._client

    def _generate(self, request: LLMRequest) -> LLMResponse:
        client = self._get_client()
        started = time.perf_counter()
        try:
            raw = client.messages.create(**build_payload(self.profile, request))
        except Exception as exc:
            # Chỉ ghi tên lỗi: thông báo gốc có thể lặp lại nội dung yêu cầu.
            raise LLMUnavailable(f"Không gọi được model: {type(exc).__name__}") from exc
        latency_ms = int((time.perf_counter() - started) * 1000)

        if raw.stop_reason == "refusal":
            raise LLMRefusal(str(raw.stop_details))
        if raw.stop_reason == "max_tokens":
            raise LLMTruncated("Phản hồi bị cắt do max_tokens")

        texts: list[str] = []
        calls: list[ToolCall] = []
        for block in raw.content:
            if block.type == "text":
                texts.append(block.text)
            elif block.type == "tool_use":
                if not isinstance(block.input, dict):
                    raise LLMBadOutput("Đối số tool không phải object")
                calls.append(ToolCall(block.id, block.name, block.input))

        return LLMResponse(
            provider=self.profile.provider,
            model=raw.model,
            text="".join(texts),
            tool_calls=tuple(calls),
            request_id=raw.id,
            input_tokens=raw.usage.input_tokens,
            output_tokens=raw.usage.output_tokens,
            latency_ms=latency_ms,
        )
