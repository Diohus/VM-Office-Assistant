"""Trần chi phí: code chặn trước khi gọi, theo số token và số tiền mỗi lần chạy (Q55).

Không có số dùng thì tính theo mức xấu nhất. Không có giá mà có trần tiền thì không gọi.
"""
from __future__ import annotations

from dataclasses import dataclass, field

from .base import ImagePart, LLMError, LLMProvider, LLMRequest, LLMResponse, TextPart

CHARS_PER_TOKEN_LOW = 2.5  # ước lượng thấp để chặn sớm; tiếng Việt có dấu tốn nhiều token hơn tiếng Anh
TOKENS_PER_IMAGE = 3000  # giả định, phải đo khi có key


class BudgetExceeded(LLMError):
    pass


@dataclass(frozen=True)
class Price:
    input_per_mtok: float
    output_per_mtok: float
    as_of: str
    source: str


@dataclass(frozen=True)
class Budget:
    max_tokens_per_run: int
    max_usd_per_run: float | None = None
    price: Price | None = None
    allow_missing_price: bool = False  # chỉ để chạy xác minh model; trần token vẫn áp dụng


def estimate_input_tokens(request: LLMRequest) -> int:
    chars = len(request.system)
    images = 0
    for message in request.messages:
        for part in message.parts:
            if isinstance(part, TextPart):
                chars += len(part.text)
            elif isinstance(part, ImagePart):
                images += 1
    return int(chars / CHARS_PER_TOKEN_LOW) + images * TOKENS_PER_IMAGE


def _usd(price: Price | None, input_tokens: int, output_tokens: int) -> float | None:
    if price is None:
        return None
    return input_tokens / 1_000_000 * price.input_per_mtok + output_tokens / 1_000_000 * price.output_per_mtok


@dataclass
class BudgetedProvider:
    """Bọc một provider. Một đối tượng là một lần chạy, nên tạo mới cho mỗi lần chạy."""

    inner: LLMProvider
    budget: Budget
    tokens_used: int = 0
    usd_used: float = 0.0
    calls: list[LLMResponse] = field(default_factory=list)

    @property
    def profile(self):
        return self.inner.profile

    @property
    def is_fake(self) -> bool:
        return self.inner.is_fake

    def check_capabilities(self, request: LLMRequest) -> None:
        self.inner.check_capabilities(request)

    def generate(self, request: LLMRequest) -> LLMResponse:
        budget = self.budget
        # Khả năng kiểm tra trước, để model thiếu khả năng báo đúng lỗi của nó.
        self.inner.check_capabilities(request)
        worst_input = estimate_input_tokens(request)
        worst_total = worst_input + request.max_output_tokens
        if self.tokens_used + worst_total > budget.max_tokens_per_run:
            raise BudgetExceeded(f"Vượt trần {budget.max_tokens_per_run} token mỗi lần chạy. Không gọi model.")
        if budget.max_usd_per_run is not None:
            worst_usd = _usd(budget.price, worst_input, request.max_output_tokens)
            if worst_usd is None and not budget.allow_missing_price:
                raise BudgetExceeded("Có trần tiền nhưng chưa có giá của model. Không gọi model.")
            if worst_usd is not None and self.usd_used + worst_usd > budget.max_usd_per_run:
                raise BudgetExceeded(f"Vượt trần {budget.max_usd_per_run} USD mỗi lần chạy. Không gọi model.")

        response = self.inner.generate(request)

        # Không có số dùng thì giữ mức xấu nhất đã ước lượng, không tính là 0.
        input_tokens = response.input_tokens if response.input_tokens is not None else worst_input
        output_tokens = response.output_tokens if response.output_tokens is not None else request.max_output_tokens
        self.tokens_used += input_tokens + output_tokens
        cost = response.cost_usd if response.cost_usd is not None else _usd(budget.price, input_tokens, output_tokens)
        if cost is not None:
            self.usd_used += cost
        self.calls.append(response)
        return response
