"""Hồ sơ biểu mẫu (Q11, Q12): khai báo cấu trúc do người xác nhận, không chứa số liệu. Backend lưu, vmai chỉ đọc."""
from __future__ import annotations

from enum import Enum

from pydantic import BaseModel, ConfigDict, Field


class ProfileStatus(str, Enum):
    NONE = "none"
    UNCONFIRMED = "unconfirmed"
    INVALIDATED = "invalidated"
    ACTIVE = "active"


class TemplateProfile(BaseModel):
    model_config = ConfigDict(frozen=True, extra="forbid")

    profile_id: str
    version: int
    structure_fingerprint: str
    source_headers: tuple[str, ...]
    column_map: dict[str, str] = Field(default_factory=dict)
    dim_maps: dict[str, dict[str, str]] = Field(default_factory=dict)
    confirmed_by: str | None = None
    confirmed_at: str | None = None


def profile_status(profile: TemplateProfile | None, fingerprint: str, source_headers: list[str]) -> ProfileStatus:
    if profile is None:
        return ProfileStatus.NONE
    if not profile.confirmed_by or not profile.confirmed_at:
        return ProfileStatus.UNCONFIRMED
    if profile.structure_fingerprint != fingerprint:
        return ProfileStatus.INVALIDATED
    if not set(source_headers) <= set(profile.source_headers):
        return ProfileStatus.INVALIDATED
    return ProfileStatus.ACTIVE
