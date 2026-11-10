"""Ca đã biết đang lọt xanh sai. Mỗi ca là xfail(strict=True): khi đã chặn được, test báo XPASS và hỏng, buộc phải bỏ dấu.

Q118 (T-AI-13 b): biểu mẫu khai sẵn GIA-QC-01, hồ sơ đã xác nhận, nạp file của máy khác cùng tên cột, cùng tháng.
Đo 10/11/2026: 180/180 ô sai bị xanh khi lệch 0,5-2%. Chưa đóng Q118 thì không phát hành, không trình diễn ô xanh (Q118).
"""
from __future__ import annotations

import importlib.util
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[2]
X01 = ROOT / "fixtures" / "eval" / "generated" / "xlsx" / "X01_m0"


def _leak(tmp_path) -> dict:
    spec = importlib.util.spec_from_file_location("run_xlsx", ROOT / "vmai" / "eval" / "run_xlsx.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module.known_leak_q118(tmp_path)


def test_q118_case_really_runs(tmp_path):
    """Không phải xfail: lỗi dựng ca (thiếu hàm, sai file) phải làm hỏng test thật, không được trốn sau xfail."""
    leak = _leak(tmp_path)
    assert leak["proposals"] == 180


@pytest.mark.xfail(strict=True, raises=AssertionError, reason="Q118: file máy khác cùng tên cột lên xanh sai; cách chặn chờ câu 47 của Owner")
def test_q118_other_machine_file_never_goes_green(tmp_path):
    assert _leak(tmp_path)["wrong_green"] == 0
