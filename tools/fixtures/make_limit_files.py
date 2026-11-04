"""Sinh 3 file giả vượt từng giới hạn của cổng đầu vào, vào fixtures/generated/ (không vào git, không vào bản chụp).

Chạy từ gốc repo bằng môi trường của reader: workers/reader/.venv/Scripts/python.exe tools/fixtures/make_limit_files.py
Sinh lại được, seed cố định.
"""
from __future__ import annotations

import io
import pathlib
import random
import zipfile

import openpyxl

ROOT = pathlib.Path(__file__).resolve().parents[2]
OUT = ROOT / "fixtures" / "generated"
QC = ROOT / "fixtures" / "templates" / "GIA-QC-01_mau.xlsx"


def with_extra_part(name: str, payload: bytes, method: int) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(QC) as zi, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for n in zi.namelist():
            zo.writestr(n, zi.read(n))
        zo.writestr(zipfile.ZipInfo(name), payload, compress_type=method)
    return out.getvalue()


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    # Trên 25 MB nhưng dưới ngưỡng chặn ở cửa: dữ liệu ngẫu nhiên, không nén được, nên số trên đĩa thật sự lớn.
    noise = random.Random(25).randbytes(int(25.5 * 1024 * 1024))
    (OUT / "GIA-VUOT_25MB.xlsx").write_bytes(with_extra_part("xl/media/dem.bin", noise, zipfile.ZIP_STORED))
    # Nhỏ trên đĩa, trên 200 MB khi giải nén.
    (OUT / "GIA-VUOT_giai_nen_200MB.xlsx").write_bytes(with_extra_part("xl/media/dem.bin", b"\0" * (205 * 1024 * 1024), zipfile.ZIP_DEFLATED))
    # Trên 200.000 ô có dữ liệu.
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet("Lon")
    ws.append(["[MẪU GIẢ] File quá nhiều ô"])
    for r in range(20_001):
        ws.append([r] * 10)
    wb.save(OUT / "GIA-VUOT_200000_o.xlsx")
    for p in sorted(OUT.iterdir()):
        print(p.name, round(p.stat().st_size / 1024 / 1024, 1), "MB")


if __name__ == "__main__":
    main()
