"""Sinh các bố cục biểu mẫu giả để đo infer_spec (Q53). Chạy: vmai/.venv/Scripts/python.exe fixtures/eval/generate_layouts.py

[MẪU GIẢ] Bố cục do chính người viết infer_spec tự nghĩ, nên đoán đúng ở đây KHÔNG chứng minh được với mẫu thật.
Đáp án lấy từ dữ liệu sinh ra.
"""
from __future__ import annotations

import datetime as dt
import json
import random
import shutil
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter as L

OUT = Path(__file__).resolve().parent / "generated" / "layouts"
TESTS = [("GLU", "Glucose", "1", 5.50, 0.15), ("GLU", "Glucose", "2", 16.50, 0.40), ("CHOL", "Cholesterol", "1", 3.00, 0.08), ("CHOL", "Cholesterol", "2", 6.50, 0.15)]
DAYS = [dt.date(2026, 9, d) for d in range(1, 11)]
ITEMS = [("VT001", "Ống nghiệm giả", "Cái"), ("VT002", "Đầu côn giả", "Hộp"), ("VT003", "Hóa chất giả A", "Lọ"), ("VT004", "Hóa chất giả B", "Lọ"), ("VT005", "Găng tay giả", "Đôi")]
DISCLAIMER = "[MẪU GIẢ] Bố cục tự nghĩ, không phải biểu mẫu thật."


def qc_values(seed: int) -> dict:
    rng = random.Random(seed)
    return {(d, code, level): round(rng.gauss(target, sd * 0.7), 2) for d in DAYS for code, _, level, target, sd in TESTS}


def long_source(path: Path, values: dict, *, extra=None, headers=("Ngày chạy", "Mã XN", "Mức", "Kết quả", "Đơn vị")) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "KetQua"
    extra = extra or {}
    ws.append([*headers, *extra])
    for (d, code, level), value in values.items():
        ws.append([d, code, f"L{level}", value, "mmol/L", *extra.values()])
    for row in ws.iter_rows(min_row=2, max_col=1):
        row[0].number_format = "dd/mm/yyyy"
    wb.properties.title = "[MẪU GIẢ] Nguồn"
    wb.save(path)


def save_case(name: str, wb, source_writer, *, cells: dict, expect: str, note: str, forbidden=()) -> None:
    directory = OUT / name
    directory.mkdir(parents=True, exist_ok=True)
    wb.properties.title = "[MẪU GIẢ]"
    wb.save(directory / "template.xlsx")
    source_writer(directory / "source.xlsx")
    (directory / "answer.json").write_text(
        json.dumps(
            {"case": name, "disclaimer": DISCLAIMER, "note": note, "expect": expect, "cells": {k: str(v) for k, v in cells.items()}, "forbidden": list(forbidden)},
            ensure_ascii=False,
            indent=1,
        ),
        encoding="utf-8",
    )


def title(ws, text="[MẪU GIẢ] BIỂU MẪU GIẢ ĐỊNH", span="A1:F1"):
    ws.merge_cells(span)
    ws[span.split(":")[0]] = text


def limits_sheet(wb, tests=TESTS):
    ws = wb.create_sheet("GioiHan")
    ws.append(["Mã", "Mức", "Đích", "SD", "Đơn vị", "Lô"])
    for code, _, level, target, sd in tests:
        ws.append([code, int(level), target, sd, "mmol/L", f"GIA-L{level}-001"])


def l01(values):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "QC"
    title(ws)
    ws["A2"], ws["B2"] = "Mã biểu mẫu:", "GIA-X"
    ws["A4"], ws["B4"] = "Tháng:", DAYS[0]
    ws.merge_cells("A6:A7")
    ws["A6"] = "Ngày"
    ws.merge_cells("B6:C6")
    ws["B6"] = "Glucose (mmol/L)"
    ws["B7"], ws["C7"] = "Mức 1", "Mức 2"
    ws.merge_cells("D6:E6")
    ws["D6"] = "Cholesterol (mmol/L)"
    ws["D7"], ws["E7"] = "Mức 1", "Mức 2"
    ws["A8"] = "=$B$4"
    for i in range(1, 10):
        ws[f"A{8 + i}"] = f"=A{7 + i}+1"
    limits_sheet(wb)
    letters = {("GLU", "1"): "B", ("GLU", "2"): "C", ("CHOL", "1"): "D", ("CHOL", "2"): "E"}
    cells = {f"{letters[(c, l)]}{8 + d.day - 1}": v for (d, c, l), v in values.items()}
    save_case("L01_hai_dong_tieu_de", wb, lambda p: long_source(p, values), cells=cells, expect="ok", note="Như GIA-QC-01: tiêu đề hai dòng, cột ngày là công thức")


def _wide_one_row(ws, top: int, left: int, values, *, date_label="Ngày"):
    headers = [date_label, "GLU M1", "GLU M2", "CHOL M1", "CHOL M2"]
    for i, text in enumerate(headers):
        ws.cell(row=top, column=left + i, value=text)
    for i, d in enumerate(DAYS):
        ws.cell(row=top + 1 + i, column=left, value=d).number_format = "dd/mm/yyyy"
    keys = [("GLU", "1"), ("GLU", "2"), ("CHOL", "1"), ("CHOL", "2")]
    return {f"{L(left + 1 + keys.index((c, l)))}{top + 1 + DAYS.index(d)}": v for (d, c, l), v in values.items()}


def l02(values):
    wb = openpyxl.Workbook()
    ws = wb.active
    title(ws)
    cells = _wide_one_row(ws, 3, 1, values)
    save_case("L02_mot_dong_tieu_de", wb, lambda p: long_source(p, values), cells=cells, expect="ok", note="Tiêu đề một dòng, viết tắt GLU M1")


def l03(values):
    wb = openpyxl.Workbook()
    ws = wb.active
    title(ws)
    ws.merge_cells("A3:A5")
    ws["A3"] = "Ngày"
    ws.merge_cells("B3:E3")
    ws["B3"] = "Sinh hóa"
    ws.merge_cells("B4:C4")
    ws["B4"] = "Glucose"
    ws.merge_cells("D4:E4")
    ws["D4"] = "Cholesterol"
    for c, text in zip("BCDE", ["Mức 1", "Mức 2", "Mức 1", "Mức 2"]):
        ws[f"{c}5"] = text
    for i, d in enumerate(DAYS):
        ws[f"A{6 + i}"] = d
    letters = {("GLU", "1"): "B", ("GLU", "2"): "C", ("CHOL", "1"): "D", ("CHOL", "2"): "E"}
    cells = {f"{letters[(c, l)]}{6 + d.day - 1}": v for (d, c, l), v in values.items()}
    save_case("L03_ba_dong_tieu_de", wb, lambda p: long_source(p, values), cells=cells, expect="ok", note="Tiêu đề ba dòng, có nhóm trên cùng")


def l04(values):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["B2"] = "[MẪU GIẢ] PHIẾU GIẢ ĐỊNH"
    ws["B4"], ws["C4"] = "Khoa:", "Khoa giả định"
    ws["B5"], ws["C5"] = "Thiết bị:", "Máy giả"
    ws["B6"], ws["C6"] = "Phiên bản:", "1.0"
    cells = _wide_one_row(ws, 10, 3, values)
    save_case("L04_bang_lech_vi_tri", wb, lambda p: long_source(p, values), cells=cells, expect="ok", note="Bảng bắt đầu ở C10, phía trên có các cặp nhãn và giá trị")


def _items_table(ws, top, *, stt=False):
    headers = (["STT"] if stt else []) + ["Mã", "Tên", "ĐVT", "Số lượng", "Đơn giá", "Thành tiền"]
    for i, text in enumerate(headers):
        ws.cell(row=top, column=1 + i, value=text)
    offset = 1 if stt else 0
    for r, (code, name, unit) in enumerate(ITEMS):
        row = top + 1 + r
        if stt:
            ws.cell(row=row, column=1, value=r + 1)
        ws.cell(row=row, column=1 + offset, value=code)
        ws.cell(row=row, column=2 + offset, value=name)
        ws.cell(row=row, column=3 + offset, value=unit)
        ws.cell(row=row, column=6 + offset, value=f"={L(4 + offset)}{row}*{L(5 + offset)}{row}")
    return offset


def _item_values(seed):
    rng = random.Random(seed)
    return {code: (rng.randint(5, 300), rng.randint(5, 900) * 100) for code, _, _ in ITEMS}


def _item_source(values, headers=("Mã VT", "Số lượng", "Đơn giá", "Lý do")):
    def write(path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(list(headers))
        for code, (qty, price) in values.items():
            ws.append([code, qty, price, "xuất dùng"])
        wb.save(path)

    return write


def _item_cells(values, top, offset):
    cells = {}
    for r, (code, _, _) in enumerate(ITEMS):
        qty, price = values[code]
        cells[f"{L(4 + offset)}{top + 1 + r}"] = qty
        cells[f"{L(5 + offset)}{top + 1 + r}"] = price
    return cells


def l05():
    wb = openpyxl.Workbook()
    ws = wb.active
    title(ws)
    offset = _items_table(ws, 3)
    values = _item_values(5)
    save_case("L05_danh_sach_vat_tu", wb, _item_source(values), cells=_item_cells(values, 3, offset), expect="ok", note="Khóa là mã vật tư, có cột công thức Thành tiền, nguồn dạng bảng cùng tiêu đề")


def l06():
    wb = openpyxl.Workbook()
    ws = wb.active
    title(ws)
    offset = _items_table(ws, 3, stt=True)
    end = 3 + len(ITEMS) + 1
    ws[f"B{end}"] = "Tổng cộng"
    ws[f"G{end}"] = f"=SUM(G4:G{end - 1})"
    ws[f"A{end + 2}"], ws[f"D{end + 2}"] = "Người lập", "Người kiểm tra"
    values = _item_values(6)
    save_case("L06_co_stt_dong_tong", wb, _item_source(values), cells=_item_cells(values, 3, offset), expect="ok", note="Cột STT đứng đầu nhưng không phải khóa, dòng tổng, ô ký")


def l07():
    wb = openpyxl.Workbook()
    ws = wb.active
    title(ws)
    ws.append([])
    for i, text in enumerate(["Ngày", "Ca", "GLU", "CHOL"]):
        ws.cell(row=3, column=1 + i, value=text)
    rng = random.Random(7)
    rows, cells = [], {}
    row = 4
    data = {}
    for d in DAYS[:5]:
        ws.merge_cells(f"A{row}:A{row + 1}")
        ws[f"A{row}"] = d
        for shift in ("Sáng", "Chiều"):
            ws[f"B{row}"] = shift
            for letter, code, target in (("C", "GLU", 5.5), ("D", "CHOL", 3.0)):
                value = round(rng.gauss(target, 0.1), 2)
                data[(d, shift, code)] = value
                cells[f"{letter}{row}"] = value
            rows.append(row)
            row += 1

    def write(path):
        src = openpyxl.Workbook()
        s = src.active
        s.append(["Ngày", "Ca", "XN", "KQ"])
        for (d, shift, code), value in data.items():
            s.append([d, shift, code, value])
        src.save(path)

    save_case("L07_khoa_hai_cot", wb, write, cells=cells, expect="ok", note="Khóa gồm ngày gộp ô và ca")


def l08(values):
    wb = openpyxl.Workbook()
    ws = wb.active
    title(ws)
    ws["A3"], ws["B3"] = "Xét nghiệm", "Mức"
    for i, d in enumerate(DAYS):
        ws.cell(row=3, column=3 + i, value=d)
    cells = {}
    for r, (code, _, level, _, _) in enumerate(TESTS):
        ws[f"A{4 + r}"], ws[f"B{4 + r}"] = code, f"L{level}"
        for i, d in enumerate(DAYS):
            cells[f"{L(3 + i)}{4 + r}"] = values[(d, code, level)]
    save_case("L08_ngang_doc_dao", wb, lambda p: long_source(p, values), cells=cells, expect="ok", note="Ngày nằm ngang, xét nghiệm nằm dọc")


def l09(values):
    wb = openpyxl.Workbook()
    ws = wb.active
    title(ws)
    cells = _wide_one_row(ws, 3, 1, values)
    _items_table(ws, 18)
    save_case("L09_hai_bang_mot_sheet", wb, lambda p: long_source(p, values), cells=cells, expect="ok_or_ask", note="Hai bảng trong một sheet, nguồn chỉ khớp bảng trên")


def l10(values):
    wb = openpyxl.Workbook()
    guide = wb.active
    guide.title = "Huong_dan"
    guide["A1"] = "[MẪU GIẢ] Hướng dẫn điền"
    guide["A2"] = "Điền đủ các ô, không sửa công thức."
    ws = wb.create_sheet("So_lieu")
    cells = _wide_one_row(ws, 2, 1, values)
    limits_sheet(wb)
    save_case("L10_nhieu_sheet", wb, lambda p: long_source(p, values), cells=cells, expect="ok", note="Sheet đầu là hướng dẫn, bảng ở sheet thứ hai, có sheet giới hạn")


def l11(values):
    wb = openpyxl.Workbook()
    ws = wb.active
    title(ws)
    ws.merge_cells("A3:A5")
    ws["A3"] = "Ngày"
    ws.merge_cells("B3:C3")
    ws["B3"] = "Glucose"
    ws.merge_cells("D3:E3")
    ws["D3"] = "Cholesterol"
    ws.merge_cells("B4:C4")
    ws["B4"] = "mmol/L"
    ws.merge_cells("D4:E4")
    ws["D4"] = "mmol/L"
    for c, text in zip("BCDE", ["Mức 1", "Mức 2", "Mức 1", "Mức 2"]):
        ws[f"{c}5"] = text
    for i, d in enumerate(DAYS):
        ws[f"A{6 + i}"] = d
    letters = {("GLU", "1"): "B", ("GLU", "2"): "C", ("CHOL", "1"): "D", ("CHOL", "2"): "E"}
    cells = {f"{letters[(c, l)]}{6 + d.day - 1}": v for (d, c, l), v in values.items()}
    save_case("L11_don_vi_rieng_dong", wb, lambda p: long_source(p, values), cells=cells, expect="ok", note="Đơn vị nằm ở một dòng tiêu đề riêng")


def l12(values):
    wb = openpyxl.Workbook()
    ws = wb.active
    title(ws)
    for i, text in enumerate(["Ngày", "Đường huyết thấp", "Đường huyết cao", "Mỡ máu thấp", "Mỡ máu cao"]):
        ws.cell(row=3, column=1 + i, value=text)
    for i, d in enumerate(DAYS):
        ws[f"A{4 + i}"] = d
    save_case("L12_ten_khac_han", wb, lambda p: long_source(p, values), cells={}, expect="ask", note="Tên cột không có gì để khớp với GLU, L1: phải hỏi")


def l13():
    wb = openpyxl.Workbook()
    ws = wb.active
    title(ws)
    for i, text in enumerate(["Mã", "Tên", "Số lượng"]):
        ws.cell(row=3, column=1 + i, value=text)
    for r, code in enumerate(["VT001", "VT002", "VT001", "VT003"]):
        ws[f"A{4 + r}"], ws[f"B{4 + r}"] = code, f"Vật tư giả {r}"

    def write(path):
        src = openpyxl.Workbook()
        s = src.active
        s.append(["Mã", "Số lượng"])
        for code, qty in (("VT001", 5), ("VT002", 6), ("VT003", 7)):
            s.append([code, qty])
        src.save(path)

    save_case("L13_khoa_trung", wb, write, cells={}, expect="ask", note="Mã lặp trong biểu mẫu: không biết điền dòng nào")


def l14(values):
    wb = openpyxl.Workbook()
    ws = wb.active
    title(ws)
    ws["A3"], ws["B3"] = "Tên máy:", None
    ws["A4"], ws["B4"] = "Ngày nhập:", None
    ws["A5"], ws["B5"] = "Hãng:", None
    ws["A7"] = "Nhận xét chung của khoa về thiết bị."
    save_case("L14_khong_co_bang", wb, lambda p: long_source(p, values), cells={}, expect="ask", note="Biểu mẫu dạng điền nhãn, không có bảng")


def l15(values):
    wb = openpyxl.Workbook()
    ws = wb.active
    title(ws)
    for i, text in enumerate(["Ngày", "GLU M1", "GLU M2", "CHOL M1", "CHOL M2", "Người chạy", "Người kiểm tra", "Ghi chú"]):
        ws.cell(row=3, column=1 + i, value=text)
    for i, d in enumerate(DAYS):
        ws[f"A{4 + i}"] = d
    keys = [("GLU", "1"), ("GLU", "2"), ("CHOL", "1"), ("CHOL", "2")]
    cells = {f"{L(2 + keys.index((c, l)))}{4 + d.day - 1}": v for (d, c, l), v in values.items()}
    extra = {"Người chạy": "Lê Giả A", "Họ tên người bệnh": "Nguyễn Giả Một"}
    save_case("L15_cot_nguoi", wb, lambda p: long_source(p, values, extra=extra), cells=cells, expect="ok", note="Có cột người và ghi chú: không được điền", forbidden=["Lê Giả A", "Nguyễn Giả Một"])


def l16():
    wb = openpyxl.Workbook()
    ws = wb.active
    title(ws)
    for i, text in enumerate(["Mã", "Tên", "Số lượng", "Đơn giá"]):
        ws.cell(row=3, column=1 + i, value=text)
    values = _item_values(16)
    row = 4
    cells = {}
    groups = {"Vật tư": ITEMS[:2], "Hóa chất": ITEMS[2:]}
    for group, items in groups.items():
        ws.merge_cells(f"A{row}:D{row}")
        ws[f"A{row}"] = group
        row += 1
        for code, name, _ in items:
            ws[f"A{row}"], ws[f"B{row}"] = code, name
            cells[f"C{row}"], cells[f"D{row}"] = values[code]
            row += 1
    save_case("L16_dong_nhom_xen_giua", wb, _item_source(values), cells=cells, expect="ok", note="Dòng tên nhóm gộp ô xen giữa dữ liệu")


def l17(values):
    wb = openpyxl.Workbook()
    ws = wb.active
    title(ws)
    for i, text in enumerate(["Ngày", "Kết quả", "Kết quả"]):
        ws.cell(row=3, column=1 + i, value=text)
    for i, d in enumerate(DAYS):
        ws[f"A{4 + i}"] = d
    save_case("L17_hai_cot_cung_ten", wb, lambda p: long_source(p, values), cells={}, expect="ask", note="Hai cột cùng tên: không biết cột nào là số nào")


def l18(values):
    wb = openpyxl.Workbook()
    ws = wb.active
    title(ws)
    ws.merge_cells("A3:A4")
    ws["A3"] = "Ngày"
    ws.merge_cells("B3:C3")
    ws["B3"] = "Glucose"
    ws["B4"], ws["C4"] = "Thấp", "Cao"
    ws.merge_cells("D3:E3")
    ws["D3"] = "Cholesterol"
    ws["D4"], ws["E4"] = "Thấp", "Cao"
    for i, d in enumerate(DAYS):
        ws[f"A{5 + i}"] = d
    save_case("L18_muc_bang_chu", wb, lambda p: long_source(p, values), cells={}, expect="ask", note="Thấp, Cao so với L1, L2: đoán theo thứ tự là nguy hiểm, phải hỏi")


# ---------- Đợt 2: cố tình gài bẫy cho infer_spec, viết sau khi xem kết quả đợt đầu ----------
def _item_template(top=3, headers=("STT", "Mã", "Tên", "Số lượng")):
    wb = openpyxl.Workbook()
    ws = wb.active
    title(ws)
    for i, text in enumerate(headers):
        ws.cell(row=top, column=1 + i, value=text)
    return wb, ws


def l19():
    """STT 1..5 trùng với cột Kho của nguồn, trong khi khóa thật là Mã."""
    wb, ws = _item_template()
    cells = {}
    rng = random.Random(19)
    data = []
    for r, (code, name, _) in enumerate(ITEMS):
        ws[f"A{4 + r}"], ws[f"B{4 + r}"], ws[f"C{4 + r}"] = r + 1, code, name
        qty = rng.randint(5, 90)
        cells[f"D{4 + r}"] = qty
        data.append((5 - r, code, qty))

    def write(path):
        src = openpyxl.Workbook()
        s = src.active
        s.append(["Kho", "Mã", "Số lượng"])
        for row in data:
            s.append(list(row))
        src.save(path)

    save_case("L19_stt_trung_cot_khac", wb, write, cells=cells, expect="ok_or_ask", note="Cột STT có giá trị trùng cột Kho của nguồn: dễ nhận nhầm khóa")


def l20(values):
    """Nguồn có hai cột ngày cùng khớp các ngày trong biểu mẫu."""
    wb = openpyxl.Workbook()
    ws = wb.active
    title(ws)
    cells = _wide_one_row(ws, 3, 1, values)

    def write(path):
        src = openpyxl.Workbook()
        s = src.active
        s.append(["Ngày nhận mẫu", "Ngày chạy", "Mã XN", "Mức", "Kết quả"])
        for (d, code, level), value in values.items():
            received = d - dt.timedelta(days=1) if d.day > 1 else d
            s.append([received, d, code, f"L{level}", value])
        src.save(path)

    save_case("L20_hai_cot_ngay", wb, write, cells=cells, expect="ask", note="Hai cột ngày trong nguồn: không biết cột nào là khóa, phải hỏi")


def l21(values):
    """Số lưu dạng chữ, kiểu rất hay gặp."""
    wb = openpyxl.Workbook()
    ws = wb.active
    title(ws)
    cells = _wide_one_row(ws, 3, 1, values)

    def write(path):
        src = openpyxl.Workbook()
        s = src.active
        s.append(["Ngày chạy", "Mã XN", "Mức", "Kết quả"])
        for (d, code, level), value in values.items():
            s.append([d, code, f"L{level}", f"{value:.2f}".replace(".", ",")])
        src.save(path)

    save_case("L21_so_dang_chu", wb, write, cells=cells, expect="ok", note="Cột kết quả của nguồn là số dạng chữ, dấu phẩy thập phân")


def l22(values):
    """Ngày trong tháng ghi bằng số 1..10 ngang đầu bảng."""
    wb = openpyxl.Workbook()
    ws = wb.active
    title(ws)
    ws["A3"], ws["B3"] = "Xét nghiệm", "Mức"
    for i, d in enumerate(DAYS):
        ws.cell(row=3, column=3 + i, value=d.day)
    cells = {}
    for r, (code, _, level, _, _) in enumerate(TESTS):
        ws[f"A{4 + r}"], ws[f"B{4 + r}"] = code, f"L{level}"
        for i, d in enumerate(DAYS):
            cells[f"{L(3 + i)}{4 + r}"] = values[(d, code, level)]
    save_case("L22_ngay_la_so", wb, lambda p: long_source(p, values), cells=cells, expect="ok_or_ask", note="Đầu bảng là số 1..10, cách ghi nhật ký tháng hay gặp")


def l23(values):
    """Mức trong nguồn là số 1, 2 chứ không phải L1, L2."""
    wb = openpyxl.Workbook()
    ws = wb.active
    title(ws)
    cells = _wide_one_row(ws, 3, 1, values)

    def write(path):
        src = openpyxl.Workbook()
        s = src.active
        s.append(["Ngày chạy", "Mã XN", "Mức", "Kết quả"])
        for (d, code, level), value in values.items():
            s.append([d, code, int(level), value])
        src.save(path)

    save_case("L23_muc_la_so", wb, write, cells=cells, expect="ok", note="Cột Mức của nguồn là số, dễ bị coi là cột kết quả")


def l24(values):
    """M1, M2 trong biểu mẫu nghĩa là Máy 1, Máy 2, nhưng nguồn có L1, L2 nghĩa là mức."""
    wb = openpyxl.Workbook()
    ws = wb.active
    title(ws)
    _wide_one_row(ws, 3, 1, values)
    ws["A2"] = "Ghi chú đầu trang: M1 là máy số 1, M2 là máy số 2"
    save_case("L24_m1_nghia_khac", wb, lambda p: long_source(p, values), cells={}, expect="ask", note="M1 nghĩa là máy, không đọc được từ cấu trúc. Code không thể biết: bẫy này đo đường đoán sai không báo")


def l25(values):
    """Biểu mẫu mỗi xét nghiệm một cột, nguồn có hai mức: thiếu một chiều."""
    wb = openpyxl.Workbook()
    ws = wb.active
    title(ws)
    for i, text in enumerate(["Ngày", "Glucose", "Cholesterol"]):
        ws.cell(row=3, column=1 + i, value=text)
    for i, d in enumerate(DAYS):
        ws[f"A{4 + i}"] = d
    save_case("L25_nguon_them_chieu", wb, lambda p: long_source(p, values), cells={}, expect="ask", note="Nguồn có hai mức cho mỗi ngày mà biểu mẫu chỉ có một cột: không biết lấy mức nào")


def l26(values):
    """Có cột trống ngăn giữa hai nhóm cột."""
    wb = openpyxl.Workbook()
    ws = wb.active
    title(ws, span="A1:G1")
    for c, text in zip("ABCEF", ["Ngày", "GLU M1", "GLU M2", "CHOL M1", "CHOL M2"]):
        ws[f"{c}3"] = text
    for i, d in enumerate(DAYS):
        ws[f"A{4 + i}"] = d
    letters = {("GLU", "1"): "B", ("GLU", "2"): "C", ("CHOL", "1"): "E", ("CHOL", "2"): "F"}
    cells = {f"{letters[(c, l)]}{4 + d.day - 1}": v for (d, c, l), v in values.items()}
    save_case("L26_cot_trong_ngan", wb, lambda p: long_source(p, values), cells=cells, expect="ok", note="Một cột trống ngăn giữa hai nhóm cột")


def l27(values):
    """Tiêu đề có xuống dòng và khoảng trắng thừa, nguồn viết hoa."""
    wb = openpyxl.Workbook()
    ws = wb.active
    title(ws)
    for i, text in enumerate([" Ngày ", "GLU\nM1", "GLU\nM2", " CHOL  M1", "CHOL M2 "]):
        ws.cell(row=3, column=1 + i, value=text)
    for i, d in enumerate(DAYS):
        ws[f"A{4 + i}"] = d
    keys = [("GLU", "1"), ("GLU", "2"), ("CHOL", "1"), ("CHOL", "2")]
    cells = {f"{L(2 + keys.index((c, l)))}{4 + d.day - 1}": v for (d, c, l), v in values.items()}
    save_case("L27_tieu_de_ban", wb, lambda p: long_source(p, values, headers=("NGÀY CHẠY", "MÃ XN", "MỨC", "KẾT QUẢ", "ĐƠN VỊ")), cells=cells, expect="ok", note="Tiêu đề có xuống dòng, khoảng trắng; nguồn viết hoa")


def l28(values):
    """Cột khóa ngày có giá trị đã lưu sẵn, nhưng tháng nguồn khác."""
    other = {(d.replace(month=10), c, l): v for (d, c, l), v in values.items()}
    wb = openpyxl.Workbook()
    ws = wb.active
    title(ws)
    _wide_one_row(ws, 3, 1, values)
    save_case("L28_sai_thang", wb, lambda p: long_source(p, other), cells={}, expect="ask", note="Biểu mẫu tháng 9, nguồn tháng 10: không có ngày nào trùng")


def _swapped_levels_case(name: str, tests, note: str, out_of_control=(), seed: int | None = None) -> None:
    """T-AI-12: nguồn của hãng đánh số mức ngược (L1 của nguồn là mức cao). App hỏi "L1 có phải M1 không",
    người dùng bấm nhầm "Cùng một thứ", nên mọi số vào nhầm mức. Có sheet giới hạn, để Q17 là thứ duy nhất chặn xanh.

    out_of_control: (ngày, mã, mức, giá trị) là lần chạy QC hỏng thật, cố ý rơi vào miền giới hạn của mức kia.
    """
    rng = random.Random(1200 + len(name) if seed is None else seed)
    true = {(d, code, level): round(rng.gauss(target, sd * 0.7), 2) for d in DAYS for code, _, level, target, sd in tests}
    for d, code, level, value in out_of_control:
        true[(d, code, level)] = value
    source = {(d, code, "2" if level == "1" else "1"): v for (d, code, level), v in true.items()}
    wb = openpyxl.Workbook()
    ws = wb.active
    title(ws)
    _wide_one_row(ws, 3, 1, true)
    limits_sheet(wb, tests)
    # cells rỗng: cách ghép sau câu trả lời là sai, nên mọi đề xuất là sai; xanh nào cũng là sai bị xanh.
    save_case(name, wb, lambda p: long_source(p, source), cells={}, expect="ask", note=note)


def l29():
    tests = [("GLU", "Glucose", "1", 5.50, 0.10), ("GLU", "Glucose", "2", 6.20, 0.10), ("CHOL", "Cholesterol", "1", 3.00, 0.05), ("CHOL", "Cholesterol", "2", 3.40, 0.05)]
    _swapped_levels_case("L29_bam_nham_muc_gan_nhau", tests, "Bấm nhầm Cùng một thứ, hai mức gần nhau, miền ±3SD không chồng (5,2-5,8 và 5,9-6,5)")


def l30():
    tests = [("GLU", "Glucose", "1", 5.50, 0.15), ("GLU", "Glucose", "2", 5.80, 0.15), ("CHOL", "Cholesterol", "1", 3.00, 0.08), ("CHOL", "Cholesterol", "2", 3.40, 0.08)]
    _swapped_levels_case("L30_bam_nham_muc_chong_mot_phan", tests, "Bấm nhầm Cùng một thứ, miền ±3SD của hai mức chồng một phần")


def l31():
    tests = [("GLU", "Glucose", "1", 5.50, 0.10), ("GLU", "Glucose", "2", 6.11, 0.10), ("CHOL", "Cholesterol", "1", 3.00, 0.05), ("CHOL", "Cholesterol", "2", 3.31, 0.05)]
    bad_runs = [(DAYS[0], "GLU", "2", 5.75), (DAYS[1], "GLU", "2", 5.70), (DAYS[0], "CHOL", "2", 3.12), (DAYS[1], "CHOL", "2", 3.10)]
    _swapped_levels_case("L31_bam_nham_muc_sat_nhau_qc_hong", tests,
                         "Bấm nhầm Cùng một thứ, miền ±3SD sát nhau không chồng, có 4 lần QC hỏng của mức 2 rơi vào miền mức 1", bad_runs)


def l32():
    """Như L31 nhưng chỉ 3 ngày: cột quá ít ô nên CHECK_COLUMN_SUSPECT không bật. Trước Q114 có 4 ô sai bị xanh ở lần sau."""
    global DAYS
    full = DAYS
    DAYS = full[:3]
    try:
        tests = [("GLU", "Glucose", "1", 5.50, 0.10), ("GLU", "Glucose", "2", 6.11, 0.10), ("CHOL", "Cholesterol", "1", 3.00, 0.05), ("CHOL", "Cholesterol", "2", 3.31, 0.05)]
        bad_runs = [(DAYS[0], "GLU", "2", 5.75), (DAYS[1], "GLU", "2", 5.70), (DAYS[0], "CHOL", "2", 3.12), (DAYS[1], "CHOL", "2", 3.10)]
        _swapped_levels_case("L32_bam_nham_muc_lo_3_ngay", tests,
                             "Bấm nhầm Cùng một thứ, miền ±3SD sát nhau, 2/3 ngày QC mức 2 hỏng rơi vào miền mức 1, lô 3 ngày", bad_runs,
                             seed=1200 + len("L31_bam_nham_muc_sat_nhau_qc_hong"))  # cùng số với lần thử phá đã báo PM
    finally:
        DAYS = full


def _wide_machine_case(name: str, headers, right: str, expect: str, note: str) -> None:
    """T-AI-13 (a): biểu mẫu ngang Ngày, Natri, Kali; nguồn ngang có thêm cột của máy thứ hai lệch khoảng 1%."""
    rng = random.Random(1300)
    wb = openpyxl.Workbook()
    ws = wb.active
    title(ws)
    for i, text in enumerate(("Ngày", "Natri", "Kali")):
        ws.cell(row=3, column=1 + i, value=text)
    rows, cells = [], {}
    for i, d in enumerate(DAYS):
        ws.cell(row=4 + i, column=1, value=d).number_format = "dd/mm/yyyy"
        na1 = round(rng.gauss(140, 1), 1)
        na2 = round(na1 * (1 + rng.uniform(-0.01, 0.01)), 1)
        na2 = na2 if na2 != na1 else round(na1 + 0.1, 1)
        k = round(rng.gauss(4.0, 0.05), 2)
        values = {"Natri máy 1": na1, "Natri": na1 if right == "Natri" else na2, "Natri máy 2": na2, "Kali": k}
        if right == "Natri máy 1":
            values["Natri"] = na2
        rows.append([d, *(values[h] for h in headers[1:])])
        cells[f"B{4 + i}"], cells[f"C{4 + i}"] = na1, k

    def source(path):
        src = openpyxl.Workbook()
        sheet = src.active
        sheet.append(list(headers))
        for row in rows:
            sheet.append(row)
        for row in sheet.iter_rows(min_row=2, max_col=1):
            row[0].number_format = "dd/mm/yyyy"
        src.properties.title = "[MẪU GIẢ] Nguồn"
        src.save(path)

    save_case(name, wb, source, cells=cells, expect=expect, note=note)


def l33():
    _wide_machine_case("L33_ngang_may_1_ten_khac", ("Ngày", "Natri máy 1", "Natri", "Kali"), "Natri máy 1", "ask",
                       "Nguồn ngang: máy 1 là 'Natri máy 1', máy 2 mang tên trần 'Natri'. Không được tự chọn cột tên trần mà không hỏi")


def l34():
    _wide_machine_case("L34_ngang_may_2_ten_khac", ("Ngày", "Natri", "Natri máy 2", "Kali"), "Natri", "ok_or_ask",
                       "Nguồn ngang: 'Natri' là máy 1, có thêm 'Natri máy 2'. App không biết cột nào đúng, hỏi lại cũng được")


def main() -> None:
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True)
    values = qc_values(1)
    for build in (l01, l02, l03, l04, l08, l09, l10, l11, l12, l14, l15, l17, l18):
        build(values)
    for build in (l05, l06, l07, l13, l16):
        build()
    for build in (l20, l21, l22, l23, l24, l25, l26, l27, l28):
        build(values)
    l19()
    for build in (l29, l30, l31, l32, l33, l34):
        build()
    print("Bố cục:", len(list(OUT.iterdir())))


if __name__ == "__main__":
    import sys

    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    main()
