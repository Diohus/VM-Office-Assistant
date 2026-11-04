"""Adapter giả để chạy không cần mạng. Kết quả của adapter này không phải kết quả chất lượng."""
from __future__ import annotations

from collections.abc import Callable

from .base import LLMProvider, LLMRequest, LLMResponse, Profile


class FakeProvider(LLMProvider):
    is_fake = True

    def __init__(
        self,
        profile: Profile,
        allow_unverified: bool = False,
        handler: Callable[[LLMRequest], LLMResponse] | None = None,
    ):
        super().__init__(profile, allow_unverified)
        self.handler = handler
        self.requests: list[LLMRequest] = []

    def _generate(self, request: LLMRequest) -> LLMResponse:
        self.requests.append(request)
        if self.handler is not None:
            return self.handler(request)
        return LLMResponse(provider=self.profile.provider, model=self.profile.model, text="{}")
