"""Xem LibreOffice ghi tên vùng gì, thử sửa đúng tên đó, rồi so bản đã sửa với .xls gốc trên Excel thật.

Chạy: python legacy_fix.py <ảnh docker>. Kết quả ghi vào results/legacy_fix.json.
"""
from __future__ import annotations

import io
import json
import pathlib
import re
import sys
import tempfile
import zipfile

from excel import excel
from legacy_check import FIX, OUT, convert, diff, facts

NAME = re.compile(rb"<definedName\b[^>]*>.*?</definedName>|<definedName\b[^>]*/>", re.S)


def only_names(data: bytes, keep) -> bytes:
    zi = zipfile.ZipFile(io.BytesIO(data))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for n in zi.namelist():
            d = zi.read(n)
            if n == "xl/workbook.xml":
                d = NAME.sub(lambda m: m.group(0) if keep(m.group(0)) else b"", d)
                d = re.sub(rb"<definedNames>\s*</definedNames>", b"", d)
            zo.writestr(n, d)
    return out.getvalue()


def main() -> int:
    image = sys.argv[1]
    res: dict = {}
    with tempfile.TemporaryDirectory(prefix="vm_fix_", ignore_cleanup_errors=True) as tmp:
        tmp = pathlib.Path(tmp)
        (tmp / "in").mkdir()
        outdir = tmp / "out"
        outdir.mkdir()
        outdir.chmod(0o777)
        x = tmp / "in" / "x.xls"
        x.write_bytes((FIX / "GIA-QC-01_da_dien_doi_cu.xls").read_bytes())
        convert(image, x, "xlsx:Calc MS Excel 2007 XML", outdir)
        data = (outdir / "x.xlsx").read_bytes()
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            names = [m.decode("utf8") for m in NAME.findall(z.read("xl/workbook.xml"))]
        res["names"] = names
        with excel(limit_seconds=240) as xl:
            orig = xl.work.copy_in(x, "goc.xls")
            wa, _ = xl.open(orig)
            a = facts(wa)
            res["names_in_xls"] = [[n.Name, n.RefersTo] for n in wa.Names]
            wa.Close(False)
            res["bad"] = []
            for i, name in enumerate(names):
                raw = name.encode("utf8")
                p = xl.work.write(f"chi_bo_{i}.xlsx", only_names(data, lambda m, raw=raw: m != raw))
                wb, err = xl.open(p)
                if wb is not None:
                    wb.Close(False)
                res["bad"].append({"bo": name, "mo_duoc": err is None})
            removable = [n.encode("utf8") for n, r in zip(names, res["bad"]) if r["mo_duoc"]]
            p = xl.work.write("da_sua.xlsx", only_names(data, lambda m: m not in removable))
            wb, err = xl.open(p)
            res["fixed_open_error"] = err
            if wb is not None:
                res["fixed_names"] = [[n.Name, n.RefersTo] for n in wb.Names]
                res["fixed_diff"] = diff(a, facts(wb))
                wb.Close(False)
    (OUT / "legacy_fix.json").write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf8")
    print(json.dumps(res, ensure_ascii=False, indent=1)[:6000])
    return 0


if __name__ == "__main__":
    sys.exit(main())
