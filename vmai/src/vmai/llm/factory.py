"""Chọn provider và model chỉ bằng cấu hình."""
from __future__ import annotations

import os
from pathlib import Path
from typing import Any

import yaml

from .base import Capability, CapabilityStatus, ConfigError, LLMProvider, Profile
from .budget import Budget, BudgetedProvider, Price

DEFAULT_CONFIG = Path(__file__).resolve().parents[1] / "config" / "llm.yaml"
PROFILE_ENV = "VMAI_LLM_PROFILE"


def _status(value: Any, where: str) -> CapabilityStatus:
    try:
        return CapabilityStatus(value)
    except ValueError as exc:
        raise ConfigError(f"{where}: giá trị khả năng không hợp lệ: {value!r}") from exc


def _budget(raw: Any) -> Budget | None:
    if not raw:
        return None
    price = raw.get("price")
    return Budget(
        max_tokens_per_run=int(raw["max_tokens_per_run"]),
        max_usd_per_run=float(raw["max_usd_per_run"]) if raw.get("max_usd_per_run") is not None else None,
        price=Price(float(price["input_per_mtok"]), float(price["output_per_mtok"]), str(price["as_of"]), str(price["source"])) if price else None,
    )


def load_profile(name: str | None = None, path: Path | str = DEFAULT_CONFIG) -> Profile:
    data = yaml.safe_load(Path(path).read_text(encoding="utf-8"))
    chosen = name or os.environ.get(PROFILE_ENV) or data.get("active")
    profiles = data.get("profiles", {})
    if chosen not in profiles:
        raise ConfigError(f"Không có profile {chosen!r} trong {path}. Có: {sorted(profiles)}")
    raw = profiles[chosen]
    for required in ("provider", "model"):
        if not raw.get(required):
            raise ConfigError(f"Profile {chosen}: thiếu {required}")
    capabilities = {
        Capability(k): _status(v, f"Profile {chosen}.{k}") for k, v in (raw.get("capabilities") or {}).items()
    }
    return Profile(
        chosen,
        raw["provider"],
        raw["model"],
        raw.get("api_key_env"),
        raw.get("base_url"),
        float(raw.get("timeout_seconds", 60)),
        capabilities,
        raw.get("routing") or {},
        _budget(raw.get("budget")),
    )


def create_provider(profile: Profile, allow_unverified: bool = False, **kwargs: Any) -> LLMProvider | BudgetedProvider:
    """Provider thật luôn được bọc trần chi phí. Không khai báo trần thì không tạo được."""
    if profile.provider == "fake":
        from .fake import FakeProvider

        return FakeProvider(profile, allow_unverified, **kwargs)
    if profile.budget is None:
        raise ConfigError(f"Profile {profile.name}: thiếu budget. Không gọi model khi không có trần chi phí.")
    if profile.provider == "openai":
        from .openai_adapter import OpenAIProvider

        inner: LLMProvider = OpenAIProvider(profile, allow_unverified, **kwargs)
    elif profile.provider == "anthropic":
        from .anthropic_adapter import AnthropicProvider

        inner = AnthropicProvider(profile, allow_unverified, **kwargs)
    elif profile.provider == "openrouter":
        from .openrouter_adapter import OpenRouterProvider

        inner = OpenRouterProvider(profile, allow_unverified, **kwargs)
    else:
        raise ConfigError(f"Không biết provider {profile.provider!r}")
    return BudgetedProvider(inner, profile.budget)
