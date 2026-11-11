"""T-DEV-06b: nút Hoàn tác, Làm lại không sáng khi bấm vào không làm gì. Xem browser/undo.mjs."""
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
    proc = subprocess.run([node, "undo.mjs"], cwd=HERE / "browser", capture_output=True, text=True, encoding="utf8", timeout=400, env={**os.environ, "VM_BASE": BASE})
    assert proc.returncode == 0, proc.stderr[-2500:]
    return json.loads((OUT / "undo.json").read_text(encoding="utf8"))


def test_khong_loi_trang(run):
    assert run["page_errors"] == []


def test_nut_sang_dung_luc(run):
    assert run["open"] == {"Hoàn tác": False, "Làm lại": False}, "mới mở: không có gì để làm"
    assert run["after_type"] == {"Hoàn tác": True, "Làm lại": False}
    assert run["after_undo"]["Hoàn tác"] is False and run["after_undo"]["Làm lại"] is True and run["after_undo"]["dirty"] == 0
    assert run["after_redo"]["Hoàn tác"] is True and run["after_redo"]["Làm lại"] is False and run["after_redo"]["dirty"] == 1


def test_sau_khi_luu_khong_con_lam_lai(run):
    saved = run["after_save"]
    assert saved["state"] == "Bản 2" and saved["Hoàn tác"] is False and saved["Làm lại"] is False


def test_het_50_buoc_thi_nut_hoan_tac_tat(run):
    limit = run["limit"]
    assert limit["Hoàn tác"] is False, "hết bước thì nút tắt"
    assert limit["clicks"] <= 52, "chỉ bấm được tới khi hết 50 bước"
    assert limit["dirty"] > 0, "vẫn còn sửa vượt 50 bước, nút không được giả vờ còn làm được"
