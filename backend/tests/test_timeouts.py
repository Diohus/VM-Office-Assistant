"""Hạn chờ dọc đường web (nginx) -> api -> llm -> egress -> OpenAI (PM 14/11). Tầng ngoài phải chờ lâu hơn tầng trong, để lỗi hết
giờ luôn do llm báo (CHAT_MODEL_TIMEOUT), không tầng nào cắt trước. Đọc giá trị mặc định trong mã, không cần Docker."""
from __future__ import annotations

import pathlib
import re

ROOT = pathlib.Path(__file__).resolve().parents[2]


def _default(path: str, pattern: str) -> float:
    match = re.search(pattern, (ROOT / path).read_text(encoding="utf8"))
    assert match, f"không thấy hạn chờ trong {path}"
    return float(match.group(1))


def test_thu_tu_han_cho():
    upstream = _default("workers/llm/llm_service.py", r'LLM_UPSTREAM_TIMEOUT_SECONDS", "([\d.]+)"')
    api = _default("backend/app/llm_client.py", r'LLM_TIMEOUT_SECONDS", "([\d.]+)"')
    egress_idle = _default("workers/egress/egress.py", r"IDLE_SECONDS = ([\d.]+)")
    nginx = _default("frontend/nginx.conf", r"proxy_read_timeout (\d+)s;")
    assert upstream >= 45, "PM 14/11: model có thể mất 23 giây cho 28 công thức; hạn chờ OpenAI ít nhất 45 giây"
    assert egress_idle > upstream, "egress không được cắt kết nối im lặng trước khi llm hết hạn chờ"
    assert api >= upstream + 10, "api phải chờ llm lâu hơn llm chờ OpenAI"
    assert nginx > api, "nginx không được cắt trước api"
