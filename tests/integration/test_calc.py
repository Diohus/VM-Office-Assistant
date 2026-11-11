"""T-DEV-08: ô công thức không hiện số cũ mà không báo sau khi sửa; chờ bộ tính có hạn. Xem browser/calc.mjs."""
from __future__ import annotations

import json
import os
import pathlib
import subprocess

import pytest

from conftest import BASE

HERE = pathlib.Path(__file__).parent
OUT = HERE / "out"


@pytest.fixture(scope="module")
def run():
    node = "node.exe" if os.name == "nt" else "node"
    proc = subprocess.run([node, "calc.mjs"], cwd=HERE / "browser", capture_output=True, text=True, encoding="utf8", timeout=580, env={**os.environ, "VM_BASE": BASE})
    assert proc.returncode == 0, proc.stderr[-2500:]
    return json.loads((OUT / "calc.json").read_text(encoding="utf8"))


def test_khong_loi_trang(run):
    assert run["page_errors"] == []


def test_cho_tinh_xong_moi_dung_so(run):
    big = run["big"]
    expected = round(big["j3_before"] - big["b3_before"] + 1000, 6)
    # Trước khi sửa (đo 10/11): settled() trả về khi J3 còn số cũ, khoảng 4 giây.
    assert big["settled"]["ok"] is True
    assert round(big["j3_after_settled"], 6) == expected, f"J3 phải là số mới {expected}, không phải số cũ {big['j3_before']}"


def test_so_cu_khong_hien_lau_ma_khong_bao(run):
    big = run["big"]
    # Dòng báo hiện sau 300 ms để không nhấp nháy; cộng thời gian vẽ và bước đo 20 ms.
    assert big["max_unannounced_ms"] < 400, f"số cũ hiện {big['max_unannounced_ms']} ms mà không có dòng báo"
    if big["stale_ms"] >= 600:
        assert big["status_text"] and "Đang tính lại" in big["status_text"]
    assert big["status_after"] is False, "tính xong thì dòng báo biến mất"


def test_bo_tinh_khong_xong_thi_co_han_cho_va_bao_dung_ly_do(run):
    stuck = run["stuck"]
    assert stuck["status_text"] and "Đang tính lại" in stuck["status_text"]
    assert stuck["settled"]["ok"] is False and stuck["settled"]["timedOut"] is True and stuck["settled"]["ms"] >= 2500, "settled() chờ bộ tính, quá hạn thì báo"
    assert stuck["downloaded"] is False, "không xuất PDF khi số có thể còn cũ"
    assert 2500 <= stuck["ms"] < 20000, "có hạn chờ, không chờ mãi"
    assert "chưa tính xong" in stuck["error"] and "không bị đổi" in stuck["error"] and "khác số trên bảng" not in stuck["error"]
    assert stuck["i12"] == 42, "app không treo, vẫn gõ được"
    assert stuck["status_after_release"] is False
