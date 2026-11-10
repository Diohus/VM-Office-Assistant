"""In được tiếng Việt trên cửa sổ Windows dùng bảng mã cũ (cp1258), không cần đặt PYTHONIOENCODING.

Chỉ đổi cách in ra màn hình. Không đụng logic, không đọc biến môi trường.
Gọi trong khối `if __name__ == "__main__":` của script eval, để test nạp script không bị đổi luồng in.
"""
from __future__ import annotations

import sys


def utf8() -> None:
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8", errors="replace")
