"""Khung chat lát 2 (Q135): đồng hồ, trần chi phí, bản xem trước giữ trong bộ nhớ. Không ghi nội dung vào log hay Postgres.

- Bản xem trước (`ChatPrepared`, có giá trị ô) chỉ nằm trong bộ nhớ api, 10 phút, xóa ngay khi đã gửi.
- Trần: mỗi lượt và mỗi ngày theo giờ Việt Nam (UTC+7, không đổi giờ theo mùa). So bằng ước tính xấu nhất trước khi gọi.
- Đồng hồ thử CHAT_TEST_NOW và trần khác mặc định chỉ dùng trên bản thử; khi bật thì /api/health báo chat_test_settings.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import os
import secrets
import threading
import time
import uuid
from dataclasses import dataclass, field

VN = dt.timezone(dt.timedelta(hours=7))
DEFAULT_MAX_USD_PER_TURN = 0.02
DEFAULT_MAX_USD_PER_DAY = 1.0
PREVIEW_SECONDS = 600
MAX_PREVIEWS = 200


def _float_env(name: str, default: float) -> float:
    raw = os.environ.get(name, "").strip()
    if not raw:
        return default
    value = float(raw)
    if not 0 <= value <= 100:
        raise ValueError(f"{name} ngoài khoảng cho phép")
    return value


def _test_now() -> dt.datetime | None:
    raw = os.environ.get("CHAT_TEST_NOW", "").strip()
    if not raw:
        return None
    value = dt.datetime.fromisoformat(raw)
    if value.tzinfo is None:
        raise ValueError("CHAT_TEST_NOW phải có múi giờ, ví dụ 2026-11-13T06:30:00+07:00")
    return value


@dataclass(frozen=True)
class Settings:
    max_usd_per_turn: float
    max_usd_per_day: float
    test_now: dt.datetime | None

    @property
    def is_test(self) -> bool:
        return (self.test_now is not None or self.max_usd_per_turn != DEFAULT_MAX_USD_PER_TURN
                or self.max_usd_per_day != DEFAULT_MAX_USD_PER_DAY)

    def now(self) -> dt.datetime:
        return self.test_now if self.test_now is not None else dt.datetime.now(dt.timezone.utc)


def load_settings() -> Settings:
    return Settings(_float_env("CHAT_MAX_USD_PER_TURN", DEFAULT_MAX_USD_PER_TURN),
                    _float_env("CHAT_MAX_USD_PER_DAY", DEFAULT_MAX_USD_PER_DAY), _test_now())


def vn_day(moment: dt.datetime) -> dt.date:
    return moment.astimezone(VN).date()


class BudgetRefused(Exception):
    def __init__(self, code: str):
        super().__init__(code)
        self.code = code


def check_budget(estimated: float | None, spent_today: float, reserved: float, settings: Settings) -> None:
    """Chặn trước khi gọi. Không có ước tính (chưa có giá) thì không gọi."""
    if estimated is None:
        raise BudgetRefused("CHAT_NO_PRICE")
    if estimated > settings.max_usd_per_turn:
        raise BudgetRefused("CHAT_TURN_LIMIT")
    if spent_today + reserved + estimated > settings.max_usd_per_day:
        raise BudgetRefused("CHAT_DAY_LIMIT")


# Lỗi gọi model mà có thể nhà cung cấp vẫn tính tiền (đã gửi đi, chưa rõ kết quả): ghi bằng ước tính xấu nhất.
MAYBE_BILLED = {"CHAT_MODEL_TIMEOUT", "CHAT_MODEL_TRUNCATED", "CHAT_MODEL_FAILED", "CHAT_MODEL_UNAVAILABLE", "CHAT_MODEL_REFUSED"}


def cost_of_failure(code: str, estimated: float) -> float:
    return estimated if code in MAYBE_BILLED else 0.0


@dataclass
class Preview:
    document_id: uuid.UUID
    base_version: int
    prepared: dict
    expires: float
    prompt_sha256: str = field(init=False)

    def __post_init__(self) -> None:
        self.prompt_sha256 = hashlib.sha256(self.prepared["prompt"].encode()).hexdigest()


class Previews:
    """Bản xem trước trong bộ nhớ. Một lần gửi lấy ra và xóa luôn (không gửi lại được cùng một bản)."""

    def __init__(self, clock=time.monotonic):
        self._items: dict[str, Preview] = {}
        self._lock = threading.Lock()
        self._clock = clock

    def _sweep(self) -> None:
        now = self._clock()
        for key in [k for k, p in self._items.items() if p.expires <= now]:
            del self._items[key]

    def put(self, document_id: uuid.UUID, base_version: int, prepared: dict) -> tuple[str, Preview]:
        with self._lock:
            self._sweep()
            if len(self._items) >= MAX_PREVIEWS:
                oldest = min(self._items, key=lambda k: self._items[k].expires)
                del self._items[oldest]
            key = secrets.token_urlsafe(18)
            item = Preview(document_id, base_version, prepared, self._clock() + PREVIEW_SECONDS)
            self._items[key] = item
            return key, item

    def take(self, key: str) -> Preview | None:
        with self._lock:
            self._sweep()
            return self._items.pop(key, None)

    def __len__(self) -> int:
        with self._lock:
            self._sweep()
            return len(self._items)


class Turns:
    """Mỗi tài liệu một lượt đang chờ; tiền đã giữ chỗ cho các lượt đang gọi (để hai lượt cùng lúc không cùng vượt trần ngày)."""

    def __init__(self):
        self._busy: dict[uuid.UUID, float] = {}
        self._lock = threading.Lock()

    def start(self, document_id: uuid.UUID, estimated: float, spent_today: float, settings: Settings) -> None:
        with self._lock:
            if document_id in self._busy:
                raise BudgetRefused("CHAT_BUSY")
            check_budget(estimated, spent_today, sum(self._busy.values()), settings)
            self._busy[document_id] = estimated

    def finish(self, document_id: uuid.UUID) -> None:
        with self._lock:
            self._busy.pop(document_id, None)


PREVIEWS = Previews()
TURNS = Turns()
