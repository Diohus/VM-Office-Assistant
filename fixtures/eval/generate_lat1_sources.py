"""File nguồn giả cho nghiệm thu lát 1 (T-AI-20, Q133). [MẪU GIẢ] Mọi số, tên người, mã người bệnh đều giả.

Gốc là `fixtures/templates/GIA-NGUON_2026-09.xlsx` (X01 chuẩn). Mỗi ca chỉ đổi đúng điều ca đó cần, seed cố định.
Chạy từ gốc repo: vmai/.venv/Scripts/python.exe fixtures/eval/generate_lat1_sources.py
Kết quả ở `fixtures/templates/nguon_lat1/`. README trong thư mục đó mô tả từng ca, KHÔNG có đáp án số: PO/BA tự tính.
"""
from __future__ import annotations

import datetime as dt
import random
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[2]
BASE = ROOT / "fixtures" / "templates" / "GIA-NGUON_2026-09.xlsx"
OUT = ROOT / "fixtures" / "templates" / "nguon_lat1"
SEED = 20261111


def _load():
    wb = openpyxl.load_workbook(BASE)
    ws = wb.active
    head = [c.value for c in ws[1]]
    return wb, ws, {h: i + 1 for i, h in enumerate(head)}


def _rows(ws, col):
    return [r for r in range(2, ws.max_row + 1) if ws.cell(r, col["Ngày chạy"]).value is not None]


def _find(ws, col, day: int, code: str, level: str) -> int:
    for r in _rows(ws, col):
        v = ws.cell(r, col["Ngày chạy"]).value
        if v.day == day and ws.cell(r, col["Mã XN"]).value == code and ws.cell(r, col["Mức"]).value == level:
            return r
    raise LookupError((day, code, level))


def _save(wb, name: str) -> None:
    wb.properties.creator = "VM-Office-Assistant [MẪU GIẢ]"
    wb.properties.lastModifiedBy = None
    wb.save(OUT / name)


def _renamed(code_value: str, name: str) -> None:
    wb, ws, col = _load()
    ws.cell(1, col["Kết quả"], "KQ")
    ws.cell(1, col["Mã XN"], "Tên XN")
    for r in _rows(ws, col):
        if ws.cell(r, col["Mã XN"]).value == "GLU":
            ws.cell(r, col["Mã XN"], code_value)
    _save(wb, name)


def x02(rng):
    # Tên xét nghiệm chưa khai trong mô tả (dòng 2.5, không ngõ cụt). Không chứa "Glucose", "GLU" để khỏi khớp một phần.
    _renamed("Đường máu (giả)", "X02_doi_ten_cot.xlsx")


def x02a(rng):
    # Tên xét nghiệm đã khai trong mô tả: "Đường huyết" là tên khác của GLU trong GIA-QC-01.yaml.
    _renamed("Đường huyết", "X02a_ten_da_khai.xlsx")


def x03(rng):
    wb, ws, col = _load()
    days = sorted(rng.sample(range(2, 29), 3))
    for r in sorted(_rows(ws, col), reverse=True):
        if ws.cell(r, col["Ngày chạy"]).value.day in days:
            ws.delete_rows(r)
    _save(wb, "X03_thieu_3_ngay.xlsx")


def x04(rng):
    wb, ws, col = _load()
    r = _find(ws, col, 10, "GLU", "L1")
    values = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
    values[col["Kết quả"] - 1] = round(values[col["Kết quả"] - 1] + 0.3, 2)  # lần chạy thứ hai, khác số
    ws.insert_rows(r + 1)
    for c, v in enumerate(values, start=1):
        ws.cell(r + 1, c, v)
        ws.cell(r + 1, c).number_format = ws.cell(r, c).number_format
    _save(wb, "X04_mot_ngay_hai_lan_chay.xlsx")


def x07(rng):
    wb, ws, col = _load()
    for r in _rows(ws, col):
        if ws.cell(r, col["Mã XN"]).value == "GLU" and ws.cell(r, col["Mức"]).value == "L1":
            v = ws.cell(r, col["Kết quả"]).value
            ws.cell(r, col["Kết quả"], f"{v:.2f}".replace(".", ","))  # số lưu dạng chữ, dấu phẩy thập phân
    _save(wb, "X07_so_dang_chu_dau_phay.xlsx")


def x10(rng):
    wb, ws, col = _load()
    for day, value in zip((3, 6, 9, 13, 16), (50, 55, 527, "0,527", "52,7")):
        ws.cell(_find(ws, col, day, "GLU", "L1"), col["Kết quả"], value)
    _save(wb, "X10_so_vo_ly.xlsx")


def x13(rng):
    wb, ws, col = _load()
    for r in _rows(ws, col):
        v = ws.cell(r, col["Ngày chạy"]).value
        ws.cell(r, col["Ngày chạy"], v.replace(month=10))
    _save(wb, "X13_sai_thang_10.xlsx")


def person_names(rng):
    wb, ws, col = _load()
    names = ["Nguyễn Giả Một", "Trần Giả Hai", "Lê Giả Ba"]
    for r in _rows(ws, col):
        ws.cell(r, col["Người chạy"], rng.choice(names))
    _save(wb, "nguoi_chay_ten_gia.xlsx")


def patient_columns(rng):
    wb, ws, col = _load()
    name_col, id_col = ws.max_column + 1, ws.max_column + 2
    ws.cell(1, name_col, "Họ tên người bệnh")
    ws.cell(1, id_col, "Mã BN")
    for i, r in enumerate(_rows(ws, col), start=1):
        ws.cell(r, name_col, f"Bệnh Nhân Giả {i:03d}")
        ws.cell(r, id_col, f"BN-GIA-{i:06d}")
    _save(wb, "X17_cot_nguoi_benh_gia.xlsx")


CASES = [x02, x02a, x03, x04, x07, x10, x13, person_names, patient_columns]


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    for make in CASES:
        make(random.Random(f"{SEED}-{make.__name__}"))
    print(f"{len(CASES)} file ở {OUT.name}/")


if __name__ == "__main__":
    main()
