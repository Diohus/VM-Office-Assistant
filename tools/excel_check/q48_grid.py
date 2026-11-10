"""Trước mỗi mốc (US-1.11, Q48): so số trên lưới với số Excel tính, trên file xuất từ app.

Chạy: python q48_grid.py <file xuất.xlsx> <số lưới.json>... theo từng cặp. Số lưới do tests/integration/browser/guards.mjs ghi, lấy từ lưới thật.
Ghi results/q48_grid.json, báo số ô đã đọc và ô lệch.
Mã thoát: 0 đạt, 1 có ô lệch hoặc không có ô nào để so, 2 tham số không đủ cặp (không kiểm gì, không ghi đè kết quả).
"""
from __future__ import annotations

import json
import math
import pathlib
import sys

from excel import excel

OUT = pathlib.Path(__file__).parent / "results" / "q48_grid.json"
TOL = 1e-6  # US-1.2: khớp tới 6 chữ số thập phân


def same(excel_value, grid) -> bool:
    if excel_value in (None, "") and grid in (None, ""):
        return True
    if isinstance(excel_value, bool) or isinstance(grid, bool):
        return bool(excel_value) == bool(grid)
    if isinstance(excel_value, (int, float)) and isinstance(grid, (int, float)):
        return math.isclose(excel_value, grid, abs_tol=TOL, rel_tol=TOL)
    return str(excel_value) == str(grid)


def main() -> int:
    args = sys.argv[1:]
    if not args or len(args) % 2:
        print("Chưa kiểm gì: cần từng cặp <file xuất> <số lưới.json>. Kết quả cũ không bị ghi đè. Lệnh đầy đủ ở docs/engineering/run.md.", file=sys.stderr)
        return 2
    report = {}
    bad = 0
    with excel() as xl:
        for path, values in zip(args[::2], args[1::2]):
            grid = json.loads(pathlib.Path(values).read_text(encoding="utf8"))
            copy = xl.work.copy_in(pathlib.Path(path))
            wb, err = xl.open(copy)
            if wb is None:
                report[path] = {"open_error": err}
                bad += 1
                continue
            xl.app.CalculateFull()
            diffs = []
            for g in grid:
                got = wb.Worksheets(g["sheet"]).Range(g["cell"]).Value2
                if not same(got, g["value"]):
                    diffs.append({"sheet": g["sheet"], "cell": g["cell"], "excel": got, "grid": g["value"]})
            wb.Close(False)
            report[pathlib.Path(path).name] = {"open_error": err, "cells_read": len(grid), "diff_count": len(diffs), "diffs": diffs[:50]}
            bad += len(diffs) + (1 if err else 0) + (0 if grid else 1)
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf8")
    print(json.dumps(report, ensure_ascii=False, indent=1)[:3000])
    return 1 if bad else 0


if __name__ == "__main__":
    raise SystemExit(main())
