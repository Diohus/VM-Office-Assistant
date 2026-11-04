"""Tìm phần nào trong file LibreOffice xuất làm Excel không mở được: bỏ lần lượt từng loại thành phần rồi mở lại.

Chạy: python legacy_bisect.py <ảnh docker>. Kết quả ghi vào results/legacy_bisect.json.
"""
from __future__ import annotations

import io
import json
import pathlib
import re
import sys
import tempfile
import zipfile

from lxml import etree

from excel import excel
from legacy_check import FIX, OUT, convert

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def strip(data: bytes, sheet_tags=(), wb_tags=(), drop_parts=None) -> bytes:
    zi = zipfile.ZipFile(io.BytesIO(data))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for n in zi.namelist():
            d = zi.read(n)
            if drop_parts and re.search(drop_parts, n):
                continue
            if drop_parts and n == "[Content_Types].xml":
                root = etree.fromstring(d)
                for el in list(root):
                    if el.get("PartName") and re.search(drop_parts, el.get("PartName")[1:]):
                        root.remove(el)
                d = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            if n.startswith("xl/worksheets/sheet") and sheet_tags:
                root = etree.fromstring(d)
                for t in sheet_tags:
                    for el in list(root.iter(f"{NS}{t}")):
                        el.getparent().remove(el)
                d = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            if n == "xl/workbook.xml" and wb_tags:
                root = etree.fromstring(d)
                for t in wb_tags:
                    for el in list(root.iter(f"{NS}{t}")):
                        el.getparent().remove(el)
                d = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            zo.writestr(n, d)
    return out.getvalue()


VARIANTS = {
    "nguyen_ban": {},
    "bo_cf": {"sheet_tags": ("conditionalFormatting",)},
    "bo_dv": {"sheet_tags": ("dataValidations",)},
    "bo_extlst": {"sheet_tags": ("extLst",)},
    "bo_drawing": {"sheet_tags": ("drawing",)},
    "bo_comment_vml": {"sheet_tags": ("legacyDrawing",)},
    "bo_ca_drawing_comment": {"sheet_tags": ("drawing", "legacyDrawing")},
    "bo_header_footer": {"sheet_tags": ("headerFooter",)},
    "bo_page_setup": {"sheet_tags": ("pageSetup",)},
    "bo_print_opts": {"sheet_tags": ("printOptions", "pageMargins", "pageSetup", "headerFooter")},
    "bo_ten_vung": {"wb_tags": ("definedNames",)},
    "bo_sheetviews": {"sheet_tags": ("sheetViews",)},
    "bo_cols": {"sheet_tags": ("cols",)},
    "bo_merge": {"sheet_tags": ("mergeCells",)},
}


def main() -> int:
    image = sys.argv[1]
    res: dict = {}
    with tempfile.TemporaryDirectory(prefix="vm_bisect_", ignore_cleanup_errors=True) as tmp:
        tmp = pathlib.Path(tmp)
        (tmp / "in").mkdir()
        outdir = tmp / "out"
        outdir.mkdir()
        outdir.chmod(0o777)
        x = tmp / "in" / "x.xls"
        x.write_bytes((FIX / "GIA-QC-01_da_dien_doi_cu.xls").read_bytes())
        k = tmp / "in" / "k.xlsx"
        k.write_bytes((FIX / "GIA-KHO-02_tieu_hao.xlsx").read_bytes())
        n = tmp / "in" / "n.xlsx"
        n.write_bytes((FIX / "GIA-NGUON_2026-09.xlsx").read_bytes())
        res["convert"] = convert(image, x, "xlsx:Calc MS Excel 2007 XML", outdir)
        convert(image, k, "xlsx:Calc MS Excel 2007 XML", outdir)
        convert(image, n, "xlsx:Calc MS Excel 2007 XML", outdir)
        data = (outdir / "x.xlsx").read_bytes()
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            res["parts"] = z.namelist()
        with excel(limit_seconds=240) as xl:
            for name, kw in VARIANTS.items():
                p = xl.work.write(f"{name}.xlsx", strip(data, **kw))
                wb, err = xl.open(p)
                res[name] = err[:60] if err else "mo_duoc"
                if wb is not None:
                    wb.Close(False)
            for label, f in (("kho_qua_lo", outdir / "k.xlsx"), ("nguon_qua_lo", outdir / "n.xlsx")):
                p = xl.work.copy_in(f, f"{label}.xlsx")
                wb, err = xl.open(p)
                res[label] = err[:60] if err else "mo_duoc"
                if wb is not None:
                    wb.Close(False)
    (OUT / "legacy_bisect.json").write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf8")
    print(json.dumps(res, ensure_ascii=False, indent=1))
    return 0


if __name__ == "__main__":
    sys.exit(main())
