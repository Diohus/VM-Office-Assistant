"""R-1021: lưới trắng khi máy nặng. Xem browser/grid_load.mjs. Trước khi sửa: 7/40 lần trắng; phá thử (trả cách cũ) phải ra trắng."""
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
    proc = subprocess.run([node, "grid_load.mjs"], cwd=HERE / "browser", capture_output=True, text=True, encoding="utf8", timeout=900, env={**os.environ, "VM_BASE": BASE})
    assert proc.returncode == 0, proc.stderr[-1500:]
    return json.loads((OUT / "grid_load.json").read_text(encoding="utf8"))


def test_tai_nang_moi_lan_mo_file_deu_co_luoi(run):
    rows = run["rows"]
    assert run["runs"] == 40
    failed = [r for r in rows if r.get("failed")]
    assert failed == [], failed[:3]
    blank = [r for r in rows if not r["drawn"] or r["host_children"] == 0]
    assert blank == [], f"{len(blank)}/{len(rows)} lần lưới trắng: {blank[:3]}"


def test_khong_loi_trang_khong_cau_bao_loi(run):
    assert all(not r["page_errors"] for r in run["rows"])
    assert not any(r["error_notice"] for r in run["rows"]), "không được rơi vào câu báo lỗi chung khi mở file bình thường"
