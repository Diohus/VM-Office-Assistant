"""Sinh fixture X18b: file có liên kết ngoài thật (có phần xl/externalLinks), do Excel lưu, kèm một ô chỉ lộ liên kết qua tên vùng.

Chạy: python make_external_link.py. Kết quả: fixtures/templates/GIA-X18b_lien_ket_ngoai.xlsx. Chỉ dùng dữ liệu giả.
"""
from __future__ import annotations

import io
import pathlib
import re
import sys
import zipfile

from excel import excel

ROOT = pathlib.Path(__file__).resolve().parents[2]
FIX = ROOT / "fixtures" / "templates"
OUT = FIX / "GIA-X18b_lien_ket_ngoai.xlsx"
SRC = "GIA-X18_nguon_ngoai.xlsx"


def strip_machine_paths(data: bytes) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as zi, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for n in zi.namelist():
            d = zi.read(n)
            if n.startswith("xl/externalLinks/") and n.endswith(".rels"):
                d = re.sub(rb'Target="file:///[^"]*[\\/]([^\\/"]+)"', lambda m: b'Target="' + m.group(1) + b'"', d)
            if n == "xl/workbook.xml":
                d = re.sub(rb'<mc:AlternateContent[^>]*>\s*<mc:Choice[^>]*>\s*<x15ac:absPath.*?</mc:AlternateContent>', b"", d, flags=re.S)
            zo.writestr(n, d)
    return out.getvalue()


def main() -> int:
    with excel() as xl:
        src_path = xl.work.path / SRC
        src = xl.app.Workbooks.Add()
        src.Worksheets(1).Name = "Sheet1"
        src.Worksheets(1).Range("A1").Value2 = 1.23
        src.Worksheets(1).Range("A2").Value2 = 4.56
        src.SaveAs(str(src_path), 51)
        target = xl.work.copy_in(FIX / "GIA-QC-01_da_dien.xlsx", "dich.xlsx")
        wb, err = xl.open(target)
        if err:
            raise RuntimeError(err)
        ws = wb.Worksheets("KQ_QC_Thang")
        ws.Range("I8").Formula = f"='[{SRC}]Sheet1'!$A$1"
        wb.Names.Add("NguonNgoai", f"='[{SRC}]Sheet1'!$A$2")
        ws.Range("I9").Formula = "=NguonNgoai"
        saved = xl.work.path / "ket_qua.xlsx"
        wb.SaveAs(str(saved), 51)
        wb.Close(False)
        src.Close(False)
        with zipfile.ZipFile(saved) as z:
            links = [n for n in z.namelist() if n.startswith("xl/externalLinks/")]
            rels = "".join(z.read(n).decode("utf8") for n in links if n.endswith(".rels"))
        if not links:
            raise RuntimeError("Excel không lưu phần externalLinks")
        # Excel ghi cả đường dẫn tuyệt đối của máy, có tên người dùng, vào quan hệ phụ. Fixture không được chứa nó.
        cleaned = strip_machine_paths(saved.read_bytes())
        check = xl.work.write("kiem.xlsx", cleaned)
        wb, err = xl.open(check)
        if err:
            raise RuntimeError(f"Excel không mở được bản đã làm sạch: {err}")
        wb.Close(False)
        with zipfile.ZipFile(check) as z:
            left = [n for n in z.namelist() if b"Users" in z.read(n) or b"file:///" in z.read(n)]
        if left:
            raise RuntimeError(f"Fixture vẫn còn đường dẫn máy ở: {left}")
        OUT.write_bytes(cleaned)
        print(OUT.name, links, re.findall(r'Target="([^"]+)"', rels)[-1])
    return 0


if __name__ == "__main__":
    sys.exit(main())
