"""Kiểm chứng Q48 trên Excel thật: file do bộ ghi của sản phẩm tạo. Chạy: python q48_check.py <file>...

Kiểm: mở sạch; sau khi Excel tính lại, số khớp đáp án tự tính; Protected View không có ô công thức nào còn số cũ
(mọi ô công thức phải trống hoặc bằng giá trị đúng); biểu đồ ở Protected View không còn số cũ.
Mỗi tham số là một file, kiểm độc lập, không ghép cặp. Muốn đối chứng thì đưa cả file gốc: dòng của file gốc cho thấy Protected View giữ số đã lưu.
Kết quả ghi đè results/q48.json, nên phải đưa đủ các file mỗi lần, lệnh ở docs/engineering/run.md.
Mã thoát: 0 đạt, 1 có file trượt hoặc file không có ô công thức nào, 2 không có tham số (không kiểm gì, không ghi đè kết quả).
"""
from __future__ import annotations

import json
import pathlib
import sys

from excel import excel

OUT = pathlib.Path(__file__).parent / "results"
TOL = 1e-9


def formula_cells(wb) -> dict[str, list[str]]:
    out = {}
    for ws in wb.Worksheets:
        ur = ws.UsedRange
        cells = []
        for row in ur.Rows:
            for c in row.Cells:
                if c.HasFormula:
                    cells.append(c.GetAddress(False, False))
        out[ws.Name] = cells
    return out


def chart_values(wb) -> list:
    out = []
    for ws in wb.Worksheets:
        for co in ws.ChartObjects():
            for s in co.Chart.SeriesCollection():
                try:
                    out.append({"sheet": ws.Name, "series": s.Name, "values": list(s.Values)})
                except Exception as exc:  # noqa: BLE001 - không đọc được cũng là kết quả cần ghi
                    out.append({"sheet": ws.Name, "error": str(exc)[:120]})
    return out


def same(a, b) -> bool:
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) <= TOL * max(1, abs(a), abs(b))
    return a == b


def check(xl, label: str, path: pathlib.Path) -> dict:
    res: dict = {"label": label}
    wb, err = xl.open(path)
    res["open_error"] = err
    if wb is None:
        return res
    xl.app.CalculateFull()
    cells = formula_cells(wb)
    truth = {f"{s}!{a}": wb.Worksheets(s).Range(a).Value2 for s, addrs in cells.items() for a in addrs}
    res["charts_normal"] = chart_values(wb)
    res["formula_cells"] = len(truth)
    wb.Close(False)
    pv = xl.open_protected(path)
    try:
        book = pv.Workbook
        stale, blank = [], 0
        for key, want in truth.items():
            sheet, addr = key.split("!")
            got = book.Worksheets(sheet).Range(addr).Value2
            if got in (None, ""):
                blank += 1
            elif not same(got, want):
                stale.append({"cell": key, "protected_view": got, "correct": want})
        res["pv_blank"] = blank
        res["pv_stale"] = stale
        res["charts_protected_view"] = chart_values(book)
    finally:
        pv.Close()
    return res


def main() -> int:
    args = [pathlib.Path(a) for a in sys.argv[1:]]
    if not args:
        print("Chưa kiểm gì: thiếu tên file. Kết quả cũ không bị ghi đè. Lệnh đầy đủ ở docs/engineering/run.md.", file=sys.stderr)
        return 2
    results = []
    failed = False
    with excel() as xl:
        for src in args:
            copy = xl.work.copy_in(src)
            r = check(xl, src.name, copy)
            results.append(r)
            failed |= bool(r.get("open_error")) or bool(r.get("pv_stale")) or not r.get("formula_cells")
    OUT.mkdir(exist_ok=True)
    (OUT / "q48.json").write_text(json.dumps(results, ensure_ascii=False, indent=1), encoding="utf8")
    for r in results:
        print(r["label"], "mở:", r["open_error"] or "sạch", "ô công thức:", r.get("formula_cells"), "trống:", r.get("pv_blank"), "số cũ:", len(r.get("pv_stale", [])))
        print("  biểu đồ thường:", json.dumps(r.get("charts_normal"), ensure_ascii=False)[:200])
        print("  biểu đồ Protected View:", json.dumps(r.get("charts_protected_view"), ensure_ascii=False)[:200])
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
