"""X12: để Excel thật tính và lưu ô công thức, để giá trị đã lưu giống file thực tế.

Chạy trên bản sinh ra trong generated/, không đụng file nào khác. Excel chạy phiên riêng, ẩn, tắt macro và tự thoát.
Dùng môi trường của spike Dev vì cần pywin32: spikes/editor/.venv/Scripts/python.exe fixtures/eval/resave_x12.py
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "spikes" / "editor"))

from excel_com import excel_app  # noqa: E402

XLSX_WORKBOOK = 51
directory = ROOT / "fixtures" / "eval" / "generated" / "xlsx_lat1" / "X12"
before, after = directory / "source_before_excel.xlsx", directory / "source.xlsx"

with excel_app() as xl:
    wb = xl.Workbooks.Open(str(before))
    xl.CalculateFull()
    wb.SaveAs(str(after), FileFormat=XLSX_WORKBOOK)
    wb.Close(False)
print("Đã lưu", after)
