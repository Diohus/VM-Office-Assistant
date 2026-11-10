"""Sinh fixture giả trước khi chạy test (T-AI-17, Q128).

Thiếu fixture thì chạy `fixtures/eval/make_fixtures.py` trước khi thu thập test. Sinh hỏng thì dừng cả lần chạy với mã 1:
bẫy sai xanh không được biến mất lặng lẽ vì thiếu fixture.
"""
from __future__ import annotations

import importlib.util
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[2]
# Có project-memory/ là cây làm việc đầy đủ của đội. Kho công khai không có thư mục này, cũng không có docs/.
FULL_TREE = (ROOT / "project-memory").is_dir()


def _maker():
    spec = importlib.util.spec_from_file_location("make_fixtures", ROOT / "fixtures" / "eval" / "make_fixtures.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def pytest_sessionstart(session):
    maker = _maker()
    if not maker.missing():
        return
    try:
        maker.make()
    except SystemExit as exc:
        pytest.exit(f"Không sinh được fixture giả, test bẫy không chạy được: {exc}", returncode=1)


def pytest_terminal_summary(terminalreporter, exitstatus, config):
    """Q131: in rủi ro Owner chấp nhận ở cuối mọi lần chạy, kể cả `-q`. Không tính vào số đạt."""
    leak = getattr(config, "q131_result", None)
    if leak is None:
        line = "RỦI RO OWNER CHẤP NHẬN, Q131: chưa đo trong lần chạy này (chạy tests/test_known_leaks.py)"
    else:
        line = f"RỦI RO OWNER CHẤP NHẬN, Q131: {leak['wrong_green']}/{leak['proposals']} ô xanh sai khi nạp nhầm file máy khác không ghi mã máy"
    terminalreporter.write_line(line, yellow=True, bold=True)
