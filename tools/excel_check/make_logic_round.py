"""Lát 2 (Q135): tạo file thử IF, ROUND, ABS, AND, OR, do Excel thật tính và lưu. Không phải một phần của sản phẩm.

Mục đích: chốt danh sách hàm cho phép của khung chat. Công thức AI viết do lưới (Univer) tính, nên chỉ cho hàm đã khớp Excel.
Mỗi cột là một giá trị thử ở dòng 3 (số khó làm tròn, số âm .5, trống, chữ, TRUE, số do công thức sinh). Mỗi dòng là một công thức.
Excel tính rồi lưu, nên số đã lưu là đáp án; app mở file thì so lưới với số này (compareWithSaved), như GIA-X20.
Chạy trên máy Windows có Excel: tools/excel_check/.venv/Scripts/python.exe tools/excel_check/make_logic_round.py
Ghi fixtures/templates/GIA-X21_ham_logic_lam_tron.xlsx (dữ liệu giả, không có tên người).
"""
from __future__ import annotations

import pathlib
import sys

from clean_fixture_paths import check_same, clean_all
from excel import excel

ROOT = pathlib.Path(__file__).resolve().parents[2]
OUT = ROOT / "fixtures" / "templates" / "GIA-X21_ham_logic_lam_tron.xlsx"
XL_OPENXML = 51

# Cột: (nhãn, giá trị dòng 3). Chuỗi bắt đầu bằng "=" là công thức (số thực sinh ra khi tính).
DATA = {
    "B": ("2,675", 2.675),
    "C": ("-2,5", -2.5),
    "D": ("0,5", 0.5),
    "E": ("1,005", 1.005),
    "F": ("trống", None),
    "G": ("chữ", "abc"),
    "H": ("TRUE", True),
    "I": ("0", 0),
    "J": ("-0,0001", -0.0001),
    "K": ("12345,6789", 12345.6789),
    "L": ("=0,1+0,2", "=0.1+0.2"),
    "M": ("=1,1*1,1", "=1.1*1.1"),
    "N": ("-1,005", -1.005),
    "O": ("5,55", 5.55),
}

# {x} là ô dòng 3 của cột. Gồm từng hàm riêng và cách hay gặp trong biểu mẫu QC (làm tròn trung bình, Đạt/Không đạt).
FUNCS = [
    "ROUND({x},0)", "ROUND({x},1)", "ROUND({x},2)", "ROUND({x},3)", "ROUND({x},-1)", "ROUND({x},-2)",
    "ROUND({x}*100,0)/100", "ROUND({x},2)*1",
    "ABS({x})", "ABS({x})-{x}",
    'IF({x}>0,"Đạt","Không đạt")', "IF({x},1,0)", 'IF({x}="",1,0)', 'IF({x}>=5.5,{x},"")', "IF({x}>0,{x})",
    "IF({x}<0,ABS({x}),ROUND({x},1))", 'IF(AND({x}>=0,{x}<=10),"Trong","Ngoài")', 'IF(OR({x}<0,{x}>10),"Ngoài","Trong")',
    "AND({x}>0,{x}<10)", "OR({x}>0,{x}<-1)", "AND({x})", "OR({x})", "AND({x},TRUE)", "OR({x},FALSE)",
    'IFERROR(ROUND({x},2),"")', 'IFERROR(AND({x}),"")', 'IFERROR(IF({x}>0,1,0),"")',
    "ROUND(ABS({x}),1)", "ABS(ROUND({x},0))",
]
# Công thức trên cả dòng 3 (vùng lẫn số, chữ, trống, TRUE).
ROW_FUNCS = [
    "AND(B3:O3)", "OR(B3:O3)", "AND(F3:F3)", "OR(F3:F3)", "ROUND(AVERAGE(B3:O3),2)", "ROUND(SUM(B3:O3),1)",
    'IF(COUNT(B3:O3)>5,"Đủ","Thiếu")', "ABS(MIN(B3:O3))", 'IF(AND(COUNT(B3:O3)>0,MAX(B3:O3)>100),ROUND(MAX(B3:O3),0),"")',
]
FIRST = 10  # dòng đầu của bảng hàm


def main() -> int:
    with excel(limit_seconds=180) as xl:
        wb = xl.app.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Name = "LogicLamTron"
        ws.Range("A1").Value = "Lát 2: IF, ROUND, ABS, AND, OR (dữ liệu giả)"
        ws.Range("A2").Value = "Giá trị"
        for col, (label, value) in DATA.items():
            ws.Range(f"{col}2").Value = "'" + label
            if isinstance(value, str) and value.startswith("="):
                ws.Range(f"{col}3").Formula = value
            elif value is not None:
                ws.Range(f"{col}3").Value = value
        ws.Range(f"A{FIRST - 1}").Value = "Công thức"
        for i, f in enumerate(FUNCS):
            row = FIRST + i
            ws.Range(f"A{row}").Value = "'" + f.replace("{x}", "x")
            for col in DATA:
                ws.Range(f"{col}{row}").Formula = "=" + f.replace("{x}", f"{col}3")
        start = FIRST + len(FUNCS) + 1
        for i, f in enumerate(ROW_FUNCS):
            ws.Range(f"A{start + i}").Value = "'" + f
            ws.Range(f"B{start + i}").Formula = "=" + f
        xl.app.CalculateFull()
        path = xl.work.path / OUT.name
        wb.SaveAs(str(path), XL_OPENXML)
        wb.Close(False)
        # Gỡ đường dẫn thư mục tạm và tên người dùng Office (R-1016), rồi mở lại so công thức, giá trị.
        saved = path.read_bytes()
        cleaned, removed, names = clean_all(saved)
        check_same(xl, OUT.name, saved, cleaned)
        OUT.parent.mkdir(parents=True, exist_ok=True)
        OUT.write_bytes(cleaned)
        print(f"Đã gỡ {removed} đường dẫn máy, {names} ô tên người")
    print(f"Đã ghi {OUT.relative_to(ROOT)}: {len(FUNCS)} công thức x {len(DATA)} giá trị + {len(ROW_FUNCS)} công thức trên dòng")
    return 0


if __name__ == "__main__":
    sys.exit(main())
