"""Adapter OpenRouter, qua API Chat Completions. Chưa từng gọi thật: chờ key của khách hàng.

Giữ quy tắc không tự đổi model: đúng một ID, tắt tự chọn và tắt chuyển nhà cung cấp, kiểm tra lại model trong phản hồi.
Tài liệu hãng về định tuyến: provider.allow_fallbacks, data_collection, require_parameters. Cần kiểm chứng khi có key.
"""
from __future__ import annotations

import base64
import json
import time
from typing import Any

from .base import (
    ConfigError,
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

DEFAULT_BASE_URL = "https://openrouter.ai/api/v1"
_AUTO_MODELS = {"openrouter/auto", "auto"}
# Mặc định an toàn: không chuyển nhà cung cấp khác, không dùng nơi có thể lưu dữ liệu, không bỏ qua tham số.
DEFAULT_ROUTING = {"allow_fallbacks": False, "data_collection": "deny", "require_parameters": True}


def _content(message: Message) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    for part in message.parts:
        if isinstance(part, TextPart):
            out.append({"type": "text", "text": part.text})
        elif isinstance(part, ImagePart):
            encoded = base64.b64encode(part.data).decode("ascii")
            out.append({"type": "image_url", "image_url": {"url": f"data:{part.media_type};base64,{encoded}"}})
    return out


def check_model_id(model: str) -> None:
    """Một ID cụ thể. Không cho tự chọn, không cho danh sách model."""
    if not model or model in _AUTO_MODELS or model.endswith(":auto") or "," in model:
        raise ConfigError(f"OpenRouter phải chốt đúng một ID model, không dùng tự chọn: {model!r}")


def build_payload(profile: Profile, request: LLMRequest) -> dict[str, Any]:
    check_model_id(profile.model)
    payload: dict[str, Any] = {
        "model": profile.model,
        "messages": [{"role": "system", "content": request.system}, *({"role": m.role, "content": _content(m)} for m in request.messages)],
        "max_tokens": request.max_output_tokens,
    }
    if request.tools:
        payload["tools"] = [
            {"type": "function", "function": {"name": t.name, "description": t.description, "parameters": t.parameters, "strict": True}}
            for t in request.tools
        ]
    if request.response_schema is not None:
        payload["response_format"] = {
            "type": "json_schema",
            "json_schema": {"name": request.schema_name, "schema": request.response_schema, "strict": True},
        }
    return payload


class OpenRouterProvider(LLMProvider):
    def __init__(self, profile: Profile, allow_unverified: bool = False, client: Any = None):
        super().__init__(profile, allow_unverified)
        check_model_id(profile.model)
        self._client = client

    def _get_client(self) -> Any:
        if self._client is None:
            from openai import OpenAI

            self._client = OpenAI(
                api_key=read_api_key(self.profile),
                base_url=self.profile.base_url or DEFAULT_BASE_URL,
                timeout=self.profile.timeout_seconds,
                max_retries=2,
            )
        return self._client

    def _generate(self, request: LLMRequest) -> LLMResponse:
        client = self._get_client()
        routing = {**DEFAULT_ROUTING, **self.profile.routing}
        started = time.perf_counter()
        try:
            raw = client.chat.completions.create(**build_payload(self.profile, request), extra_body={"provider": routing})
        except Exception as exc:
            # Chỉ ghi tên lỗi: thông báo gốc có thể lặp lại nội dung yêu cầu.
            raise LLMUnavailable(f"Không gọi được model: {type(exc).__name__}") from exc
        latency_ms = int((time.perf_counter() - started) * 1000)

        # Tắt tự chọn rồi vẫn kiểm tra: nhận kết quả của model khác là tự đổi model.
        if raw.model and raw.model != self.profile.model and not raw.model.startswith(f"{self.profile.model}-"):
            raise LLMBadOutput(f"Phản hồi đến từ model khác: {raw.model}")
        if not raw.choices:
            raise LLMBadOutput("Phản hồi không có nội dung")
        choice = raw.choices[0]
        if choice.finish_reason == "length":
            raise LLMTruncated("Phản hồi bị cắt do giới hạn token")
        message = choice.message
        if getattr(message, "refusal", None):
            raise LLMRefusal(str(message.refusal))
        calls: list[ToolCall] = []
        for call in message.tool_calls or []:
            try:
                calls.append(ToolCall(call.id, call.function.name, json.loads(call.function.arguments)))
            except json.JSONDecodeError as exc:
                raise LLMBadOutput(f"Đối số tool không phải JSON: {exc}") from exc
        usage = raw.usage
        reported = getattr(usage, "cost", None) if usage else None
        return LLMResponse(
            provider=self.profile.provider,
            model=raw.model or self.profile.model,
            text=message.content or "",
            tool_calls=tuple(calls),
            request_id=raw.id,
            input_tokens=usage.prompt_tokens if usage else None,
            output_tokens=usage.completion_tokens if usage else None,
            latency_ms=latency_ms,
            cost_usd=float(reported) if isinstance(reported, (int, float)) else None,
        )
