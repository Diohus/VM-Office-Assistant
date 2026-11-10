"""Công cụ kiểm mà không kiểm gì thì không được báo đạt, và không được ghi đè kết quả cũ. Không mở Excel."""
from __future__ import annotations

import os
import pathlib
import subprocess

import pytest

ROOT = pathlib.Path(__file__).resolve().parents[2]
TOOLS = ROOT / "tools" / "excel_check"
PY = TOOLS / ".venv" / "Scripts" / "python.exe"


@pytest.mark.parametrize(
    "script, args, result",
    [
        ("q48_check.py", [], "q48.json"),
        ("q48_grid.py", [], "q48_grid.json"),
        ("q48_grid.py", ["chi_mot_file.xlsx"], "q48_grid.json"),
    ],
)
def test_khong_co_gi_de_kiem_thi_thoat_ma_loi_va_khong_ghi_de(script, args, result):
    if not PY.exists():
        pytest.skip("Máy này chưa có môi trường tools/excel_check")
    path = TOOLS / "results" / result
    before = path.read_bytes() if path.exists() else None
    proc = subprocess.run([str(PY), script, *args], cwd=TOOLS, capture_output=True, text=True, encoding="utf8", timeout=60, env={**os.environ, "PYTHONIOENCODING": "utf8"})
    assert proc.returncode == 2, proc.stderr
    assert "Chưa kiểm gì" in proc.stderr
    assert (path.read_bytes() if path.exists() else None) == before
