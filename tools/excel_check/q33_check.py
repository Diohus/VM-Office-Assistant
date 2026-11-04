"""Kiểm chứng Q33 trên Excel thật: hai bộ tính lệch thì bỏ giá trị đã lưu của ô đó và bật tính lại khi mở.

Câu hỏi: file có mở sạch không; chế độ thường có tính đúng không; Protected View hiện gì ở ô đó,
và có còn hiện số sai nào không. Chạy: python q33_check.py. Kết quả ghi vào results/q33.json.
"""
from __future__ import annotations

import io
import json
import pathlib
import re
import sys
import zipfile

from lxml import etree

from excel import excel

ROOT = pathlib.Path(__file__).resolve().parents[2]
FIX = ROOT / "fixtures" / "templates"
OUT = pathlib.Path(__file__).parent / "results"
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
SHEET_PART = "xl/worksheets/sheet1.xml"     # KQ_QC_Thang
SHEET = "KQ_QC_Thang"
PARSER = etree.XMLParser(resolve_entities=False, no_network=True)


def patch(src: bytes, set_values: dict[str, float], drop_cached: list[str], full_calc: bool = True) -> bytes:
    """Sửa số trong ô, bỏ giá trị đã lưu của các ô công thức chỉ định, bật fullCalcOnLoad. Phần khác giữ nguyên."""
    zi = zipfile.ZipFile(io.BytesIO(src))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for info in zi.infolist():
            data = zi.read(info.filename)
            if info.filename == SHEET_PART:
                root = etree.fromstring(data, PARSER)
                cells = {c.get("r"): c for c in root.iter(f"{NS}c")}
                for ref, val in set_values.items():
                    c = cells[ref]
                    c.attrib.pop("t", None)
                    for child in list(c):
                        c.remove(child)
                    etree.SubElement(c, f"{NS}v").text = repr(val)
                if drop_cached == ["*"]:
                    drop_cached = [r for r, c in cells.items() if c.find(f"{NS}f") is not None]
                for ref in drop_cached:
                    c = cells[ref]
                    if c.find(f"{NS}f") is None:
                        raise ValueError(f"{ref} không phải ô công thức")
                    v = c.find(f"{NS}v")
                    if v is not None:
                        c.remove(v)
                    c.attrib.pop("t", None)
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            elif info.filename == "xl/workbook.xml" and full_calc:
                text = data.decode("utf8")
                text, n = re.subn(r"<calcPr\b([^>]*?)/>", lambda m: f'<calcPr{m.group(1)} fullCalcOnLoad="1"/>' if "fullCalcOnLoad" not in m.group(1) else m.group(0), text)
                if n != 1:
                    raise ValueError("Không tìm thấy calcPr")
                data = text.encode("utf8")
            zo.writestr(info, data)
    return out.getvalue()


def read_cells(ws, refs):
    return {r: {"value": ws.Range(r).Value2, "text": ws.Range(r).Text} for r in refs}


def main() -> int:
    OUT.mkdir(exist_ok=True)
    src = (FIX / "GIA-QC-01_da_dien.xlsx").read_bytes()
    res: dict = {}
    # B39 là trung bình B8:B37, B40 là thống kê khác cùng cột. C39 làm đối chứng: không đổi, không bỏ giá trị.
    watch = ["A9", "B8", "B38", "B39", "B40", "B41", "B42", "B43", "C39"]
    cases = {
        # Q33: ô lệch B39 bị bỏ giá trị. Các ô thống kê khác của cột B giữ số đã được hai bộ tính đồng ý.
        "q33_mot_o_lech": {"set": {"B8": 500.0}, "drop": ["B39"]},
        # Nhiều ô lệch cùng cột, gồm cả ô là gốc của công thức dùng chung.
        "q33_ca_cot_lech": {"set": {"B8": 500.0}, "drop": ["B38", "B39", "B40", "B41", "B42", "B43"]},
        # Worker lỗi hoặc chưa so xong: mọi ô công thức của sheet bị bỏ giá trị (nhánh PO/BA hỏi trong US-1.11).
        "worker_loi_bo_ca_sheet": {"set": {"B8": 500.0}, "drop": ["*"]},
        # Đối chứng âm: không bỏ giá trị. Protected View phải hiện số cũ, nếu không thì phép đo không phân biệt được gì.
        "doi_chung_khong_bo": {"set": {"B8": 500.0}, "drop": []},
    }
    with excel() as xl:
        for name, case in cases.items():
            path = xl.work.write(f"{name}.xlsx", patch(src, case["set"], case["drop"]))
            r: dict = {"dropped": case["drop"]}
            pv = xl.open_protected(path)
            try:
                r["protected_view"] = read_cells(pv.Workbook.Worksheets(SHEET), watch)
                r["protected_view_opened"] = True
            finally:
                pv.Close()
            wb, err = xl.open(path)
            r["open_error"] = err
            if wb is not None:
                ws = wb.Worksheets(SHEET)
                r["normal"] = read_cells(ws, watch)
                r["charts_kept"] = wb.Worksheets("Bieu_do").ChartObjects().Count
                vals = [ws.Range(f"B{i}").Value2 for i in range(8, 38)]
                nums = [v for v in vals if isinstance(v, (int, float))]
                r["expected_B39"] = sum(nums) / len(nums)
                # Sau khi Excel tự tính lại và lưu thành file khác, giá trị phải quay lại file.
                saved = xl.work.path / f"{name}_saved.xlsx"
                wb.SaveAs(str(saved), 51)
                wb.Close(False)
                pv2 = xl.open_protected(saved)
                try:
                    r["protected_view_after_excel_save"] = read_cells(pv2.Workbook.Worksheets(SHEET), ["B39"])
                finally:
                    pv2.Close()
            res[name] = r
        res["excel_version"] = xl.app.Version
    (OUT / "q33.json").write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf8")
    print(json.dumps(res, ensure_ascii=False, indent=1))
    return 0


if __name__ == "__main__":
    sys.exit(main())
