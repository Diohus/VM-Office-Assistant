"""Nạp đúng một biến key từ file .env ở gốc repo vào môi trường của tiến trình này. Không in, không ghi ra đâu (Q85)."""
from __future__ import annotations

import os
from pathlib import Path

ENV_FILE = Path(__file__).resolve().parents[2] / ".env"


def env_names(path: Path = ENV_FILE) -> list[str]:
    """Chỉ tên biến, không có giá trị."""
    if not path.exists():
        return []
    names = []
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            names.append(line.split("=", 1)[0].removeprefix("export ").strip())
    return names


def load_key(name: str | None, path: Path = ENV_FILE) -> bool:
    """Trả về có nạp được không. Biến đã có trong môi trường thì giữ nguyên. Không nạp biến nào khác."""
    if not name or name in os.environ or not path.exists():
        return bool(name and name in os.environ)
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        if key.removeprefix("export ").strip() == name:
            value = value.strip().strip('"').strip("'")
            if value:
                os.environ[name] = value
                return True
    return False
