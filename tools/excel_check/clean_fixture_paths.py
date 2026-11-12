"""Q73, R-1016: gỡ đường dẫn máy và tên người khỏi fixture do Excel, Word lưu, rồi mở lại bằng Excel và so công thức, giá trị với bản trước khi dọn.

Đường dẫn: chỉ sửa thẻ absPath trong xl/workbook.xml. Tên người: tác giả, người sửa cuối, công ty, người quản lý trong thông tin file
(tools/commit_scan/people_meta.py). Các phần khác giữ nguyên. File có mật khẩu và file hỏng phải sinh lại từ bản đã dọn, vì không gỡ được
từ bên ngoài; file có mật khẩu lưu với RemovePersonalInformation (chỉ áp cho workbook đó, không đổi tên người dùng của Office).
Chạy: python clean_fixture_paths.py. Không ghi đè fixture nếu Excel báo lỗi, số khác đi, hay còn tên người.
"""
from __future__ import annotations

import io
import pathlib
import re
import shutil
import sys
import tempfile
import zipfile

import olefile

from excel import excel

ROOT = pathlib.Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "tools" / "commit_scan"))
import people_meta  # noqa: E402

FIX = ROOT / "fixtures" / "templates"
ABS_PATH = re.compile(rb"<mc:AlternateContent\b[^>]*>\s*<mc:Choice\b[^>]*>\s*<x15ac:absPath\b.*?</mc:AlternateContent>", re.S)
SKIP = {"GIA-KHO-02_kho.xlsx"}
PASSWORD = "matkhau-gia"
XL_OPENXML, XL_XLS = 51, 56
SUFFIXES = {".xlsx", ".xlsm", ".xls", ".doc"}
PROPS = ("Author", "Last Author", "Company", "Manager")
# .doc: chỉ được đổi các luồng chứa thông tin file; phần nội dung (WordDocument, Data...) phải giữ nguyên từng byte.
DOC_META_STREAMS = {"\x05SummaryInformation", "\x05DocumentSummaryInformation", "0Table", "1Table"}


def clean(data: bytes) -> tuple[bytes, int]:
    """Gỡ absPath. Trả (byte mới, số đường dẫn đã gỡ)."""
    out = io.BytesIO()
    removed = 0
    with zipfile.ZipFile(io.BytesIO(data)) as zi, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for n in zi.namelist():
            d = zi.read(n)
            if n == "xl/workbook.xml":
                d, removed = ABS_PATH.subn(b"", d)
            zo.writestr(n, d)
    return out.getvalue(), removed


def clean_all(data: bytes) -> tuple[bytes, int, int]:
    """Gỡ cả đường dẫn máy và tên người. Trả (byte mới, số đường dẫn, số ô tên người)."""
    paths = 0
    if people_meta.kind(data) == "zip":
        data, paths = clean(data)
    data, names = people_meta.strip(data)
    return data, paths, names


def values(wb) -> dict:
    out = {}
    for ws in wb.Worksheets:
        ur = ws.UsedRange
        out[ws.Name] = (repr(ur.Formula), repr(ur.Value2))
    return out


def names_left(wb) -> list[str]:
    """Ô thông tin file mà Excel còn đọc được tên (không in giá trị)."""
    left = []
    for k in PROPS:
        try:
            v = str(wb.BuiltinDocumentProperties(k).Value or "")
        except Exception:  # noqa: BLE001 - thuộc tính không có trong loại file này
            continue
        if v.replace("\x00", "").strip() not in people_meta.ALLOWED:
            left.append(k)
    return left


def check_same(xl, name: str, before: bytes, after: bytes) -> None:
    """Mở hai bản bằng Excel: không lỗi, công thức và giá trị như nhau, bản sau không còn tên người."""
    suffix = pathlib.PurePath(name).suffix
    a, err_a = xl.open(xl.work.write("truoc" + suffix, before))
    b, err_b = xl.open(xl.work.write("sau" + suffix, after))
    try:
        if err_a or err_b:
            raise RuntimeError(f"{name}: Excel báo {err_a or err_b}")
        if values(a) != values(b):
            raise RuntimeError(f"{name}: công thức hoặc giá trị đổi sau khi dọn")
        if names_left(b):
            raise RuntimeError(f"{name}: Excel còn đọc được tên ở {names_left(b)}")
    finally:
        for wb in (a, b):
            if wb is not None:
                wb.Close(False)


def check_doc(name: str, before: bytes, after: bytes) -> None:
    """.doc: chỉ luồng thông tin file đổi, nội dung giữ nguyên từng byte."""
    with olefile.OleFileIO(io.BytesIO(before)) as a, olefile.OleFileIO(io.BytesIO(after)) as b:
        sa = {"/".join(s): a.openstream(s).read() for s in a.listdir()}
        sb = {"/".join(s): b.openstream(s).read() for s in b.listdir()}
    changed = {s for s in sa if sa[s] != sb.get(s)} | (set(sb) - set(sa))
    if not changed <= DOC_META_STREAMS:
        raise RuntimeError(f"{name}: đổi cả luồng nội dung {sorted(changed - DOC_META_STREAMS)}")


def save_with_password(xl, src: bytes, name: str, fmt: int) -> bytes:
    """Sinh file có mật khẩu từ bản đã dọn. Lưu ở thư mục trung tính để đường dẫn bị mã hóa trong đó không nói lên gì,
    và với RemovePersonalInformation để Excel không ghi tên người dùng vào phần mã hóa. Mở lại bằng mật khẩu để kiểm."""
    folder = pathlib.Path(tempfile.mkdtemp(prefix="gia_", dir=ROOT.anchor))
    try:
        wb, err = xl.open(xl.work.write("goc.xlsx", src))
        if err:
            raise RuntimeError(err)
        before = values(wb)
        wb.RemovePersonalInformation = True
        saved = folder / name
        wb.SaveAs(str(saved), fmt, PASSWORD)
        wb.Close(False)
        data = saved.read_bytes()
    finally:
        shutil.rmtree(folder, ignore_errors=True)
    check = xl.work.write("kiem" + pathlib.PurePath(name).suffix, data)
    wb = xl.app.Workbooks.Open(str(xl.work.require_inside(check)), 0, True, None, PASSWORD)
    try:
        if values(wb) != before:
            raise RuntimeError(f"{name}: công thức hoặc giá trị khác bản gốc")
        if names_left(wb):
            raise RuntimeError(f"{name}: Excel còn đọc được tên ở {names_left(wb)}")
    finally:
        wb.Close(False)
    if people_meta.kind(data) != "ole_ma_hoa":
        raise RuntimeError(f"{name}: Excel không mã hóa file")
    return data


def targets() -> list[pathlib.Path]:
    out = []
    for p in sorted(FIX.rglob("*")):
        if not p.is_file() or p.suffix.lower() not in SUFFIXES or p.name in SKIP or p.name.startswith("~$"):
            continue
        if people_meta.kind(p.read_bytes()) in ("zip", "ole"):  # file mã hóa, file hỏng: sinh lại ở dưới
            out.append(p)
    return out


def main() -> int:
    done = []
    with excel() as xl:
        for path in targets():
            before = path.read_bytes()
            after, paths, names = clean_all(before)
            if not paths and not names:
                continue
            if path.suffix.lower() == ".doc":
                check_doc(path.name, before, after)
            else:
                check_same(xl, path.name, before, after)
            if [w for w, v in people_meta.people(after) if v not in people_meta.ALLOWED]:
                raise RuntimeError(f"{path.name}: còn tên người sau khi gỡ")
            path.write_bytes(after)
            done.append(f"{path.relative_to(FIX).as_posix()} ({paths} đường dẫn, {names} ô tên)")
        src = (FIX / "GIA-QC-01_da_dien.xlsx").read_bytes()
        for name, fmt in (("GIA-X19_co_mat_khau.xlsx", XL_OPENXML), ("GIA-X19_xls_mat_khau.xls", XL_XLS)):
            (FIX / name).write_bytes(save_with_password(xl, src, name, fmt))
            done.append(f"{name} (sinh lại)")
    # File hỏng là phần đầu của biểu mẫu đã dọn.
    (FIX / "GIA-X19_hong.xlsx").write_bytes((FIX / "GIA-QC-01_da_dien.xlsx").read_bytes()[:3000])
    done.append("GIA-X19_hong.xlsx (sinh lại)")
    print("Đã dọn:")
    for line in done:
        print(" ", line)
    return 0


if __name__ == "__main__":
    sys.exit(main())
