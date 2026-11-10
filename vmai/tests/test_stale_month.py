"""Q119-Q121: nghi nạp nhầm file tháng trước. Chỉ hạ màu. [MẪU GIẢ]

- Q121: ngày nào giống hệt tháng trước ở mọi cột kết quả (bỏ cột cố định) thì chính ngày đó tối đa vàng, không ngưỡng.
- Q119, Q120: từ 3 cột thay đổi, quá 30% số ngày và ít nhất 2 ngày trùng thì cả lô tối đa vàng, kèm thông báo chung.
  Ngưỡng này đặt trên phân bố tự nghĩ.
Nguyên tắc từ Q121: không có "cái giá" nào là sai xanh. Mọi ca ở đây phải có 0 ô xanh nằm ở ngày trùng.
"""
from __future__ import annotations

import importlib.util
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[2]
X01 = ROOT / "fixtures" / "eval" / "generated" / "xlsx" / "X01_m1"


@pytest.fixture(scope="module")
def cases(tmp_path_factory):
    spec = importlib.util.spec_from_file_location("run_xlsx", ROOT / "vmai" / "eval" / "run_xlsx.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return {r["case"]: r for r in module.stale_month_cases(tmp_path_factory.mktemp("cu"))}


def test_every_case_passes_its_own_check(cases):
    assert all(r["ok"] for r in cases.values()), [r["case"] for r in cases.values() if not r["ok"]]


def test_no_green_cell_on_a_day_identical_to_last_month(cases):
    assert {name: r["stale_green_cells"] for name, r in cases.items() if r["stale_green_cells"]} == {}


@pytest.mark.parametrize("edited", [0, 10, 20])
def test_reloaded_last_month_file_is_flagged_and_never_green(cases, edited):
    r = cases[f"nap_lai_thang_truoc_sua_{edited}_dong"]
    assert r["alarm"] and r["without_check"].get("green") == 180  # không có kiểm này thì 180 ô xanh
    assert "green" not in r["colors"]


def test_correct_new_month_is_untouched(cases):
    r = cases["nguon_dung_thang_moi"]
    assert not r["alarm"] and r["colors"] == {"green": 180} and r["row_yellow_cells"] == 0


@pytest.mark.parametrize("name", ["trung_tu_nhien_4_cot_co_dinh", "trung_tu_nhien_lam_tron_theo_sd", "trung_tu_nhien_3_cot_it_bien_dong"])
def test_natural_repeats_give_no_batch_alarm(cases, name):
    """Không báo chung. Ngày trùng tự nhiên vẫn vàng từng dòng: đó là cái giá của Q121, chỉ là vàng."""
    r = cases[name]
    assert not r["alarm"] and r["ok"] and r["stale_green_cells"] == 0
    # Chỉ đúng các ô của ngày trùng bị hạ: mỗi ngày X01 có 6 ô số.
    assert r["row_yellow_cells"] == 6 * r["same_rows"]
    assert r["colors"].get("yellow", 0) == r["without_check"].get("yellow", 0) + r["row_yellow_cells"]


def test_natural_repeat_cases_really_repeat(cases):
    """Các ca trùng tự nhiên phải thật sự trùng, nếu không các test trên không chứng minh gì."""
    four = cases["trung_tu_nhien_4_cot_co_dinh"]
    assert four["cells_same_share"] >= 0.5
    # Trùng cả dòng trên ngưỡng, chỉ 2 cột thay đổi: bỏ điều kiện "≤ 2 cột" của phần cả lô thì ca này báo nhầm.
    assert four["moving_columns"] == 2 and four["rows_same_share"] > 0.30
    assert four["row_yellow_cells"] > 0  # Q121 vẫn áp cho biểu mẫu chỉ 2 cột thay đổi
    assert cases["trung_tu_nhien_lam_tron_theo_sd"]["cells_same_share"] >= 0.25
    # Có dòng trùng thật nhưng dưới ngưỡng cả lô: hạ ngưỡng quá thấp thì ca này báo nhầm.
    r = cases["trung_tu_nhien_3_cot_it_bien_dong"]
    assert r["moving_columns"] == 3 and 0 < r["rows_same_share"] < 0.30 and r["row_yellow_cells"] > 0


# ---- Lô 3 ngày (Q120, Q121).


def test_short_batch_correct_source_is_untouched(cases):
    r = cases["lo_3_ngay_nguon_dung"]
    assert not r["alarm"] and r["colors"] == {"green": 18}


def test_short_batch_one_natural_same_day_gives_no_batch_alarm(cases):
    """1/3 ngày trùng là 33%, trên ngưỡng 30%. Không báo chung nhờ điều kiện ≥ 2 ngày; riêng ngày đó vàng (Q121)."""
    r = cases["lo_3_ngay_dung_1_ngay_trung"]
    assert r["rows_same_share"] > 0.30 and not r["alarm"]
    assert r["row_yellow_cells"] == 6 and r["colors"] == {"yellow": 6, "green": 12}


def test_short_batch_old_file_intact_is_flagged(cases):
    r = cases["lo_3_ngay_file_cu_nguyen"]
    assert r["alarm"] and "green" not in r["colors"]


def test_short_batch_old_file_edited_two_days_has_no_stale_green(cases):
    """Q121. Trước đây là "cái giá" của Q120: 6 ô xanh mang số tháng trước. Nay ngày không sửa tự vàng."""
    r = cases["lo_3_ngay_file_cu_sua_2_ngay"]
    assert not r["alarm"] and r["stale_green_cells"] == 0 and r["row_yellow_cells"] == 6
    assert r["colors"] == {"yellow": 6, "green": 12}
