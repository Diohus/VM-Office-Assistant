"""N9 (R-1015): tạo file thử các hàm thống kê trên vùng ít số, do Excel thật tính và lưu. Không phải một phần của sản phẩm.

Mỗi cột dữ liệu là một kiểu vùng: rỗng, 1 số, 2 số, chỉ chữ, số lẫn chữ và TRUE. Mỗi dòng là một hàm.
Excel tính rồi lưu, nên số đã lưu trong file là đáp án; app mở file thì so lưới với số này (compareWithSaved).
Chạy trên máy Windows có Excel: tools/excel_check/.venv/Scripts/python.exe tools/excel_check/make_stat_empty.py
Ghi fixtures/templates/GIA-X20_thong_ke_vung_it_so.xlsx (dữ liệu giả, không có tên người).
"""
from __future__ import annotations

import pathlib
import sys

from clean_fixture_paths import check_same, clean_all
from excel import excel

ROOT = pathlib.Path(__file__).resolve().parents[2]
OUT = ROOT / "fixtures" / "templates" / "GIA-X20_thong_ke_vung_it_so.xlsx"
XL_OPENXML = 51

# Cột dữ liệu B..G, dòng 3..7. Dòng 2 là tên kiểu vùng.
DATA = {
    "B": ("rỗng", []),
    "C": ("1 số", [5.5]),
    "D": ("2 số", [5.5, 6.1]),
    "E": ("chỉ chữ", ["abc"]),
    "F": ("số, chữ, TRUE", [5.5, "abc", True, 6.1]),
    "G": ("số âm", [-2.5, -3.5]),
}

# Không có MINIFS, MAXIFS: Excel trên máy đo không có hai hàm này (ra #NAME?), nên không có đáp án.
# Hàm, {r} là vùng. Gồm nhóm STDEV/VAR (lỗi đã biết), và các hàm thống kê hay gặp trong biểu mẫu QC, kho.
FUNCS = [
    "STDEV({r})", "STDEV.S({r})", "STDEV.P({r})", "STDEVP({r})", "STDEVA({r})", "STDEVPA({r})",
    "VAR({r})", "VAR.S({r})", "VAR.P({r})", "VARP({r})", "VARA({r})", "VARPA({r})",
    "AVERAGE({r})", "AVERAGEA({r})", 'AVERAGEIF({r},">0")', 'AVERAGEIFS({r},{r},">0")',
    "MEDIAN({r})", "MIN({r})", "MAX({r})", "MINA({r})", "MAXA({r})",
    "MODE({r})", "MODE.SNGL({r})", "GEOMEAN({r})", "HARMEAN({r})", "LARGE({r},1)", "SMALL({r},1)",
    "PERCENTILE({r},0.5)", "QUARTILE({r},1)", "KURT({r})", "SKEW({r})", "DEVSQ({r})", "AVEDEV({r})",
    "COUNT({r})", "COUNTA({r})", "SUM({r})", "PRODUCT({r})", "SUMSQ({r})",
    'IFERROR(STDEV({r}),"")', 'IFERROR(AVERAGE({r}),"")', 'IFERROR(MEDIAN({r}),"")',
]
FIRST = 10  # dòng đầu của bảng hàm


def main() -> int:
    with excel(limit_seconds=180) as xl:
        wb = xl.app.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Name = "ThongKe"
        ws.Range("A1").Value = "N9: hàm thống kê trên vùng ít số (dữ liệu giả)"
        ws.Range("A2").Value = "Kiểu vùng"
        for col, (label, values) in DATA.items():
            ws.Range(f"{col}2").Value = label
            for i, v in enumerate(values):
                ws.Range(f"{col}{3 + i}").Value = v
        ws.Range(f"A{FIRST - 1}").Value = "Hàm"
        for i, f in enumerate(FUNCS):
            row = FIRST + i
            ws.Range(f"A{row}").Value = "'" + f.replace("{r}", "vùng")
            for col in DATA:
                ws.Range(f"{col}{row}").Formula = "=" + f.replace("{r}", f"{col}3:{col}7")
        xl.app.CalculateFull()
        path = xl.work.path / OUT.name
        wb.SaveAs(str(path), XL_OPENXML)
        wb.Close(False)
        # Excel ghi đường dẫn thư mục tạm (có tên người dùng máy) vào xl/workbook.xml, và tên người dùng Office vào docProps (R-1016):
        # gỡ cả hai, rồi mở lại so công thức, giá trị và kiểm Excel không còn đọc được tên.
        saved = path.read_bytes()
        cleaned, removed, names = clean_all(saved)
        check_same(xl, OUT.name, saved, cleaned)
        OUT.parent.mkdir(parents=True, exist_ok=True)
        OUT.write_bytes(cleaned)
        print(f"Đã gỡ {removed} đường dẫn máy, {names} ô tên người")
    print(f"Đã ghi {OUT.relative_to(ROOT)}: {len(FUNCS)} hàm x {len(DATA)} kiểu vùng")
    return 0


if __name__ == "__main__":
    sys.exit(main())
