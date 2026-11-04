"""Q73: gỡ đường dẫn máy khỏi fixture do Excel lưu, rồi mở lại bằng Excel và so công thức, giá trị với bản trước khi dọn.

Chỉ sửa thẻ absPath trong xl/workbook.xml; các phần khác giữ nguyên từng byte. File có mật khẩu và file hỏng phải sinh lại từ bản đã dọn,
vì không gỡ được từ bên ngoài. Chạy: python clean_fixture_paths.py. Không ghi đè fixture nếu Excel báo lỗi hay số khác đi.
"""
from __future__ import annotations

import io
import pathlib
import re
import shutil
import sys
import tempfile
import zipfile

from excel import excel

ROOT = pathlib.Path(__file__).resolve().parents[2]
FIX = ROOT / "fixtures" / "templates"
ABS_PATH = re.compile(rb"<mc:AlternateContent\b[^>]*>\s*<mc:Choice\b[^>]*>\s*<x15ac:absPath\b.*?</mc:AlternateContent>", re.S)
SKIP = {"GIA-KHO-02_kho.xlsx"}
PASSWORD = "matkhau-gia"


def clean(data: bytes) -> tuple[bytes, int]:
    out = io.BytesIO()
    removed = 0
    with zipfile.ZipFile(io.BytesIO(data)) as zi, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for n in zi.namelist():
            d = zi.read(n)
            if n == "xl/workbook.xml":
                d, removed = ABS_PATH.subn(b"", d)
            zo.writestr(n, d)
    return out.getvalue(), removed


def values(wb) -> dict:
    out = {}
    for ws in wb.Worksheets:
        ur = ws.UsedRange
        out[ws.Name] = (repr(ur.Formula), repr(ur.Value2))
    return out


def main() -> int:
    targets = [p for p in sorted(FIX.glob("*.xls[xm]")) if p.name not in SKIP and not p.name.startswith("~$") and p.read_bytes()[:4] == b"PK\x03\x04" and zipfile.is_zipfile(p)]
    done = []
    with excel() as xl:
        for path in targets:
            before = path.read_bytes()
            cleaned, removed = clean(before)
            if not removed:
                continue
            a, err_a = xl.open(xl.work.write("truoc" + path.suffix, before))
            b, err_b = xl.open(xl.work.write("sau" + path.suffix, cleaned))
            if err_a or err_b:
                raise RuntimeError(f"{path.name}: Excel báo {err_a or err_b}")
            if values(a) != values(b):
                raise RuntimeError(f"{path.name}: công thức hoặc giá trị đổi sau khi dọn")
            a.Close(False)
            b.Close(False)
            path.write_bytes(cleaned)
            done.append(path.name)
        # File có mật khẩu: Excel mã hóa cả absPath. Sinh lại từ bản đã dọn, lưu ở thư mục trung tính để đường dẫn trong đó không nói lên gì.
        folder = pathlib.Path(tempfile.mkdtemp(prefix="gia_", dir=ROOT.anchor))
        try:
            src = xl.work.copy_in(FIX / "GIA-QC-01_da_dien.xlsx", "goc.xlsx")
            wb, err = xl.open(src)
            if err:
                raise RuntimeError(err)
            saved = folder / "GIA-X19_co_mat_khau.xlsx"
            wb.SaveAs(str(saved), 51, PASSWORD)
            wb.Close(False)
            shutil.copyfile(saved, FIX / saved.name)
            done.append(saved.name)
        finally:
            shutil.rmtree(folder, ignore_errors=True)
    # File hỏng là phần đầu của biểu mẫu đã dọn.
    (FIX / "GIA-X19_hong.xlsx").write_bytes((FIX / "GIA-QC-01_da_dien.xlsx").read_bytes()[:3000])
    done.append("GIA-X19_hong.xlsx")
    print("Đã dọn:", done)
    return 0


if __name__ == "__main__":
    sys.exit(main())
