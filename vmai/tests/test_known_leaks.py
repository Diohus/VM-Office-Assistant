"""Rủi ro Owner chấp nhận (Q131), trước là ca lọt Q118. Không xóa bẫy, không tính là đạt.

Q131: xanh nghĩa là "app chép đúng số từ file người dùng đã chọn". Chọn đúng file là việc của nhân viên.
Rủi ro: biểu mẫu khai sẵn GIA-QC-01, hồ sơ đã xác nhận, nạp nhầm file của máy khác KHÔNG ghi mã máy, cùng tên cột,
cùng tháng. Đo 10/11/2026: 180/180 ô sai bị xanh.

- `xfail(strict=True)` giữ ca này ở dạng "1 xfailed", không phải "đạt". Hết lọt thì XPASS và hỏng, buộc phải xem lại.
- Con số và phạm vi được ghim ở test thường: đổi số, hoặc file dựng ra có dấu hiệu mã máy, là lỗi.
- Cuối mọi lần chạy pytest, kể cả `-q`, `conftest.py` in dòng "RỦI RO OWNER CHẤP NHẬN, Q131".
"""
from __future__ import annotations

import importlib.util
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[2]


def _module():
    spec = importlib.util.spec_from_file_location("run_xlsx", ROOT / "vmai" / "eval" / "run_xlsx.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


@pytest.fixture(scope="module")
def leak(tmp_path_factory, pytestconfig):
    result = _module().known_leak_q118(tmp_path_factory.mktemp("q131"))
    pytestconfig.q131_result = result  # conftest in dòng rủi ro từ số đo thật của lần chạy này
    return result


def test_q131_risk_is_exactly_what_the_owner_accepted(leak):
    """Không phải xfail: dựng ca hỏng, số đổi, hay phạm vi lan sang file có mã máy đều làm hỏng test thật."""
    assert {k: leak[k] for k in _module().Q131_ACCEPTED} == _module().Q131_ACCEPTED
    assert leak["machine_marks"] == [], "Ca rủi ro chỉ gồm file KHÔNG ghi mã máy (Q131)"


def test_machine_marks_finds_a_machine_column(tmp_path):
    """Kiểm phạm vi ở trên phải thấy được mã máy khi có, nếu không thì test trên không chứng minh gì."""
    import openpyxl

    wb = openpyxl.Workbook()
    wb.active.append(["Ngày chạy", "Mã XN", "Kết quả", "Thiết bị"])
    wb.active.append(["01/09", "GLU", 5.1, "Máy giả B"])
    wb.save(tmp_path / "co_may.xlsx")
    assert _module().machine_marks(tmp_path / "co_may.xlsx")


@pytest.mark.xfail(strict=True, raises=AssertionError, reason="RỦI RO OWNER CHẤP NHẬN, Q131: nạp nhầm file máy khác không ghi mã máy thì số sai vẫn xanh")
def test_q131_other_machine_file_without_machine_code_never_goes_green(leak):
    assert leak["wrong_green"] == 0
