"""Đọc API key. Chỉ đọc đúng biến môi trường trong cấu hình, không tìm ở nơi khác."""
from __future__ import annotations

import os

from .base import MissingAPIKey, Profile


def read_api_key(profile: Profile) -> str:
    if not profile.api_key_env:
        raise MissingAPIKey(f"Profile {profile.name} chưa khai báo api_key_env")
    key = os.environ.get(profile.api_key_env)
    if not key:
        raise MissingAPIKey(f"Chưa có biến môi trường {profile.api_key_env} cho profile {profile.name}. Không gọi model.")
    return key
