"""Q48(d): biểu đồ vẽ từ ô CÔNG THỨC ở Protected View, sau khi bộ ghi bỏ số đã lưu.

Chạy: python q48_chart_formula.py. Cần môi trường có `reader`: đặt PYTHONPATH tới workers/reader và vmai/src. Kết quả: results/q48_chart_formula.json.
"""
from __future__ import annotations

import json
import pathlib
import sys

from excel import excel

ROOT = pathlib.Path(__file__).resolve().parents[2]
sys.path[:0] = [str(ROOT / "workers" / "reader"), str(ROOT / "vmai" / "src")]

from reader.writer import apply  # noqa: E402

FIX = ROOT / "fixtures" / "templates"
OUT = pathlib.Path(__file__).parent / "results"
SHEET = "KQ_QC_Thang"
LINE_CHART = 4


def series_values(book, sheet: str) -> list:
    out = []
    for co in book.Worksheets(sheet).ChartObjects():
        for s in co.Chart.SeriesCollection():
            try:
                out.append([None if v in (None, "") else float(v) for v in s.Values])
            except Exception as exc:  # noqa: BLE001 - không đọc được cũng là kết quả cần ghi
                out.append(str(exc)[:120])
    return out


def main() -> int:
    res: dict = {}
    with excel() as xl:
        wb, err = xl.open(xl.work.copy_in(FIX / "GIA-QC-01_da_dien.xlsx", "goc.xlsx"))
        if err:
            raise RuntimeError(err)
        ws = wb.Worksheets(SHEET)
        co = ws.ChartObjects().Add(700, 500, 300, 200)
        co.Chart.ChartType = LINE_CHART
        s = co.Chart.SeriesCollection().NewSeries()
        s.Values = ws.Range("B39:G39")     # thống kê: toàn ô công thức
        with_chart = xl.work.path / "co_bieu_do.xlsx"
        wb.SaveAs(str(with_chart), 51)
        wb.Close(False)
        written = apply(with_chart.read_bytes(), [{"sheet": SHEET, "cell": "B8", "value": 500.0}], "6d0c2f1a-6a3e-4a53-9e57-0f7b7f3c9a11")
        path = xl.work.write("da_ghi.xlsx", written["data"])
        res["chart_caches_cleared"] = written["chart_caches_cleared"]
        pv = xl.open_protected(path)
        try:
            res["protected_view"] = series_values(pv.Workbook, SHEET)
            res["protected_view_B39"] = pv.Workbook.Worksheets(SHEET).Range("B39").Text
        finally:
            pv.Close()
        wb, err = xl.open(path)
        res["open_error"] = err
        xl.app.CalculateFull()
        res["normal"] = series_values(wb, SHEET)
        res["normal_B39"] = wb.Worksheets(SHEET).Range("B39").Value2
        wb.Close(False)
    OUT.mkdir(exist_ok=True)
    (OUT / "q48_chart_formula.json").write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf8")
    print(json.dumps(res, ensure_ascii=False)[:1500])
    return 0


if __name__ == "__main__":
    sys.exit(main())
