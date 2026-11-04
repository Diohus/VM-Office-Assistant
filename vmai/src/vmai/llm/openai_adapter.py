"""Adapter OpenAI, dùng Responses API. Chưa từng gọi thật: chờ key của khách hàng."""
from __future__ import annotations

import base64
import json
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
            kind = "input_text" if message.role == "user" else "output_text"
            out.append({"type": kind, "text": part.text})
        elif isinstance(part, ImagePart):
            encoded = base64.b64encode(part.data).decode("ascii")
            out.append({"type": "input_image", "image_url": f"data:{part.media_type};base64,{encoded}"})
    return out


def build_payload(profile: Profile, request: LLMRequest) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "model": profile.model,
        "instructions": request.system,
        "input": [{"role": m.role, "content": _content(m)} for m in request.messages],
        "max_output_tokens": request.max_output_tokens,
        "store": False,  # không để nhà cung cấp lưu phản hồi; chính sách thực tế xem checklist
    }
    if request.tools:
        payload["tools"] = [
            {"type": "function", "name": t.name, "description": t.description, "parameters": t.parameters, "strict": True}
            for t in request.tools
        ]
    if request.response_schema is not None:
        payload["text"] = {
            "format": {"type": "json_schema", "name": request.schema_name, "schema": request.response_schema, "strict": True}
        }
    return payload


class OpenAIProvider(LLMProvider):
    def __init__(self, profile: Profile, allow_unverified: bool = False, client: Any = None):
        super().__init__(profile, allow_unverified)
        self._client = client

    def _get_client(self) -> Any:
        if self._client is None:
            from openai import OpenAI

            self._client = OpenAI(
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
            raw = client.responses.create(**build_payload(self.profile, request))
        except Exception as exc:
            # Chỉ ghi tên lỗi: thông báo gốc có thể lặp lại nội dung yêu cầu.
            raise LLMUnavailable(f"Không gọi được model: {type(exc).__name__}") from exc
        latency_ms = int((time.perf_counter() - started) * 1000)

        if raw.status == "incomplete":
            raise LLMTruncated(f"Phản hồi chưa hoàn chỉnh: {raw.incomplete_details}")
        if raw.status == "failed":
            raise LLMBadOutput(f"Yêu cầu thất bại: {raw.error}")

        texts: list[str] = []
        calls: list[ToolCall] = []
        for item in raw.output:
            if item.type == "message":
                for block in item.content:
                    if block.type == "refusal":
                        raise LLMRefusal(block.refusal)
                    if block.type == "output_text":
                        texts.append(block.text)
            elif item.type == "function_call":
                try:
                    arguments = json.loads(item.arguments)
                except json.JSONDecodeError as exc:
                    raise LLMBadOutput(f"Đối số tool không phải JSON: {exc}") from exc
                calls.append(ToolCall(item.call_id, item.name, arguments))

        usage = raw.usage
        return LLMResponse(
            provider=self.profile.provider,
            model=raw.model,
            text="".join(texts),
            tool_calls=tuple(calls),
            request_id=raw.id,
            input_tokens=usage.input_tokens if usage else None,
            output_tokens=usage.output_tokens if usage else None,
            latency_ms=latency_ms,
        )
