"""Thử chuyển .xls, .doc bằng LibreOffice trong container cách ly, rồi mở kết quả bằng Excel, Word thật và so với bản gốc.

Chạy: python legacy_check.py <ảnh docker>. Kết quả ghi vào results/legacy_<ảnh>.json. Chỉ làm trên bản sao.
"""
from __future__ import annotations

import contextlib
import json
import pathlib
import re
import subprocess
import sys
import tempfile
import time

import pythoncom
import win32com.client as win32

from excel import excel

ROOT = pathlib.Path(__file__).resolve().parents[2]
FIX = ROOT / "fixtures" / "templates"
OUT = pathlib.Path(__file__).parent / "results"
BASE = ["docker", "run", "--rm", "--network", "none", "--read-only", "--tmpfs", "/tmp", "--tmpfs", "/home/lo:uid=1000,gid=1000",
        "--cap-drop", "ALL", "--security-opt", "no-new-privileges", "--memory", "2g", "--pids-limit", "256"]


def convert(image: str, src: pathlib.Path, fmt: str, outdir: pathlib.Path) -> dict:
    t = time.perf_counter()
    cmd = BASE + ["-v", f"{src.parent}:/in:ro", "-v", f"{outdir}:/out", image, "soffice", "--headless", "--norestore", "--convert-to", fmt, "--outdir", "/out", f"/in/{src.name}"]
    p = subprocess.run(cmd, capture_output=True, text=True, timeout=240, env={**__import__("os").environ, "MSYS_NO_PATHCONV": "1"})
    return {"rc": p.returncode, "seconds": round(time.perf_counter() - t, 1), "err": p.stderr[-200:]}


def facts(wb) -> dict:
    out = {}
    for ws in wb.Worksheets:
        ur = ws.UsedRange
        cells = {}
        forms, vals = ur.Formula, ur.Value2
        if ur.Count == 1:
            forms, vals = ((forms,),), ((vals,),)
        for i, row in enumerate(forms):
            for j, f in enumerate(row):
                if f not in (None, ""):
                    v = vals[i][j]
                    cells[f"{ur.Row + i},{ur.Column + j}"] = {"f": f, "v": round(v, 9) if isinstance(v, float) else v}
        merges = set()
        for i in range(ur.Rows.Count):
            for j in range(ur.Columns.Count):
                c = ur.Cells(i + 1, j + 1)
                if c.MergeCells:
                    merges.add(c.MergeArea.Address)
        with contextlib.suppress(Exception):
            dv = ws.Cells.SpecialCells(-4174).Count
        out[ws.Name] = {"visible": ws.Visible, "cells": cells, "merges": sorted(merges), "cf": ws.Cells.FormatConditions.Count,
                        "charts": ws.ChartObjects().Count, "comments": ws.Comments.Count,
                        "fmt_B8": ws.Range("B8").NumberFormat, "width_B": ws.Columns(2).ColumnWidth}
    return out


def diff(a: dict, b: dict) -> dict:
    res = {"sheets_a": list(a), "sheets_b": list(b), "cell_diff": 0, "samples": [], "other": {}}
    for name in a:
        if name not in b:
            continue
        keys = set(a[name]["cells"]) | set(b[name]["cells"])
        for k in sorted(keys):
            x, y = a[name]["cells"].get(k), b[name]["cells"].get(k)
            if x != y:
                res["cell_diff"] += 1
                if len(res["samples"]) < 12:
                    res["samples"].append([name, k, x, y])
        for key in ("visible", "merges", "cf", "charts", "comments", "fmt_B8", "width_B"):
            if a[name][key] != b[name][key]:
                res["other"][f"{name}.{key}"] = [a[name][key], b[name][key]]
    return res


def check_doc(path: pathlib.Path) -> dict:
    """Mở .docx bằng Word thật, ở phiên riêng, không cập nhật liên kết, macro tắt. Word tắt khi xong."""
    pythoncom.CoInitialize()
    w = win32.DispatchEx("Word.Application")
    w.Visible = False
    w.DisplayAlerts = 0
    w.AutomationSecurity = 3
    res: dict = {}
    try:
        try:
            d = w.Documents.Open(str(path), False, True, False)
        except Exception as exc:  # noqa: BLE001 - lỗi mở chính là kết quả cần ghi
            return {"open_error": str(exc)[:300]}
        res = {"open_error": None, "paragraphs": d.Paragraphs.Count, "tables": d.Tables.Count, "first_style": str(d.Paragraphs(1).Style),
               "first_text": d.Paragraphs(1).Range.Text[:60], "footer": d.Sections(1).Footers(1).Range.Text[:60]}
        d.Close(0)
    finally:
        w.Quit()
        pythoncom.CoUninitialize()
    return res


def main() -> int:
    image = sys.argv[1]
    OUT.mkdir(exist_ok=True)
    res: dict = {"image": image}
    res["version"] = subprocess.run(BASE + [image, "soffice", "--version"], capture_output=True, text=True, timeout=120).stdout.strip()
    with tempfile.TemporaryDirectory(prefix="vm_legacy_", ignore_cleanup_errors=True) as tmp:
        tmp = pathlib.Path(tmp)
        indir = tmp / "in"
        outdir = tmp / "out"
        indir.mkdir()
        outdir.mkdir()
        outdir.chmod(0o777)
        xls = indir / "x.xls"
        xls.write_bytes((FIX / "GIA-QC-01_da_dien_doi_cu.xls").read_bytes())
        doc = indir / "d.doc"
        doc.write_bytes((FIX / "GIA-QT-03_quy_trinh_doi_cu.doc").read_bytes())
        res["xls_convert"] = convert(image, xls, "xlsx:Calc MS Excel 2007 XML", outdir)
        res["doc_convert"] = convert(image, doc, "docx", outdir)
        with excel() as xl:
            orig = xl.work.copy_in(xls, "goc.xls")
            wa, _ = xl.open(orig)
            a = facts(wa)
            wa.Close(False)
            conv = xl.work.copy_in(outdir / "x.xlsx", "chuyen.xlsx")
            wb, err = xl.open(conv)
            res["xls_open_error"] = err
            if wb is not None:
                res["xls_diff"] = diff(a, facts(wb))
                wb.Close(False)
            # Cách thay thế để so: chính Excel lưu .xls thành .xlsx.
            wa, _ = xl.open(orig)
            viaexcel = xl.work.path / "qua_excel.xlsx"
            wa.SaveAs(str(viaexcel), 51)
            wa.Close(False)
            wc, err = xl.open(viaexcel)
            res["excel_saveas_open_error"] = err
            res["excel_saveas_diff"] = diff(a, facts(wc))
            wc.Close(False)
        res["doc_open"] = check_doc(outdir / "d.docx")
    safe = re.sub(r"\W+", "_", image)
    (OUT / f"legacy_{safe}.json").write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf8")
    print(json.dumps(res, ensure_ascii=False, indent=1)[:5000])
    return 0


if __name__ == "__main__":
    sys.exit(main())
