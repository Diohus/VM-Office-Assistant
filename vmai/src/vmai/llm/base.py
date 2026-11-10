"""Giao diện chung của lớp LLM. Code nghiệp vụ chỉ dùng các kiểu trong file này."""
from __future__ import annotations

import json
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from enum import Enum
from typing import Any


class CapabilityStatus(str, Enum):
    YES = "yes"
    NO = "no"
    UNVERIFIED = "unverified"


class Capability(str, Enum):
    IMAGE_INPUT = "image_input"
    TOOL_CALLS = "tool_calls"
    STRUCTURED_OUTPUT = "structured_output"


class LLMError(Exception):
    pass


class ConfigError(LLMError):
    pass


class MissingAPIKey(LLMError):
    pass


class CapabilityError(LLMError):
    pass


class LLMRefusal(LLMError):
    pass


class LLMTruncated(LLMError):
    pass


class LLMBadOutput(LLMError):
    pass


class LLMUnavailable(LLMError):
    """Lỗi mạng, hết giờ, quá hạn mức, hoặc nhà cung cấp báo lỗi. Có thể cho người dùng thử lại."""


@dataclass(frozen=True)
class TextPart:
    text: str


@dataclass(frozen=True)
class ImagePart:
    data: bytes
    media_type: str = "image/png"


@dataclass(frozen=True)
class Message:
    role: str  # "user" | "assistant"
    parts: tuple[TextPart | ImagePart, ...]


@dataclass(frozen=True)
class ToolSpec:
    name: str
    description: str
    parameters: dict[str, Any]


@dataclass(frozen=True)
class LLMRequest:
    purpose: str
    system: str
    messages: tuple[Message, ...]
    tools: tuple[ToolSpec, ...] = ()
    response_schema: dict[str, Any] | None = None
    schema_name: str = "result"
    max_output_tokens: int = 16000


@dataclass(frozen=True)
class ToolCall:
    id: str
    name: str
    arguments: dict[str, Any]


@dataclass(frozen=True)
class LLMResponse:
    provider: str
    model: str
    text: str = ""
    tool_calls: tuple[ToolCall, ...] = ()
    request_id: str | None = None
    input_tokens: int | None = None
    output_tokens: int | None = None
    latency_ms: int | None = None
    cost_usd: float | None = None  # chỉ khi nhà cung cấp báo trong phản hồi
    reasoning_tokens: int | None = None  # phần token ra dùng để suy luận, nằm trong output_tokens

    def parsed(self) -> Any:
        try:
            return json.loads(self.text)
        except json.JSONDecodeError as exc:
            raise LLMBadOutput(f"Model không trả JSON hợp lệ: {exc}") from exc


@dataclass(frozen=True)
class Profile:
    name: str
    provider: str
    model: str
    api_key_env: str | None = None
    base_url: str | None = None
    timeout_seconds: float = 60.0
    capabilities: dict[Capability, CapabilityStatus] = field(default_factory=dict)
    routing: dict[str, Any] = field(default_factory=dict)
    budget: Any = None  # llm.budget.Budget; provider thật không có trần thì không tạo được


def required_capabilities(request: LLMRequest) -> set[Capability]:
    needed: set[Capability] = set()
    if request.tools:
        needed.add(Capability.TOOL_CALLS)
    if request.response_schema is not None:
        needed.add(Capability.STRUCTURED_OUTPUT)
    if any(isinstance(p, ImagePart) for m in request.messages for p in m.parts):
        needed.add(Capability.IMAGE_INPUT)
    return needed


class LLMProvider(ABC):
    is_fake = False

    def __init__(self, profile: Profile, allow_unverified: bool = False):
        self.profile = profile
        self.allow_unverified = allow_unverified

    def check_capabilities(self, request: LLMRequest) -> None:
        """Thiếu khả năng thì báo lỗi, không âm thầm đổi model hay bỏ bớt yêu cầu."""
        for cap in sorted(required_capabilities(request), key=lambda c: c.value):
            status = self.profile.capabilities.get(cap, CapabilityStatus.UNVERIFIED)
            if status is CapabilityStatus.YES:
                continue
            if status is CapabilityStatus.UNVERIFIED and self.allow_unverified:
                continue
            raise CapabilityError(
                f"Model {self.profile.model} ({self.profile.provider}): khả năng {cap.value} "
                f"đang là '{status.value}'. Xem docs/ai/gpt6_luna_checklist.md."
            )

    def generate(self, request: LLMRequest) -> LLMResponse:
        self.check_capabilities(request)
        return self._generate(request)

    @abstractmethod
    def _generate(self, request: LLMRequest) -> LLMResponse: ...
