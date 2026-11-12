"""R-1016: tên người trong thông tin file Office (tác giả, người sửa cuối, công ty, người quản lý). Đọc để commit_scan báo, gỡ để làm sạch file mẫu.

Excel, Word tự ghi tên người dùng máy vào file khi lưu, ở những chỗ không thấy trên trang tính:
- .xlsx/.xlsm/.docx: docProps/core.xml (dc:creator, cp:lastModifiedBy), docProps/app.xml (Company, Manager).
- .xls/.doc: luồng SummaryInformation (tác giả, người lưu cuối), DocumentSummaryInformation (công ty, người quản lý).
- .xls: bản ghi WRITEACCESS trong luồng Workbook (người lưu). .doc: bảng SttbfAssoc trong luồng 0Table/1Table (tác giả, người sửa cuối).
Gỡ là ghi đè đúng chỗ đó bằng chuỗi rỗng (zip) hoặc byte 0/dấu cách cùng độ dài (file đời cũ), các phần khác giữ nguyên.
File có mật khẩu: phần này nằm trong gói mã hóa, không đọc hay gỡ được từ bên ngoài; phải sinh lại (clean_fixture_paths.py).
Không in giá trị tìm được ra log: chỉ báo vị trí.
"""
from __future__ import annotations

import io
import os
import re
import struct
import tempfile
import zipfile

import olefile

# Giá trị do công cụ hoặc đội đặt, không phải tên người. Thêm vào đây phải được PM đồng ý.
ALLOWED = {"", "openpyxl", "VM-Office-Assistant", "VM-Office-Assistant [MẪU GIẢ]"}
_ZIP_FIELDS = {
    "docProps/core.xml": (rb"dc:creator", rb"cp:lastModifiedBy"),
    "docProps/app.xml": (rb"Company", rb"Manager"),
}
_SUMMARY = "\x05SummaryInformation"
_DOC_SUMMARY = "\x05DocumentSummaryInformation"
_OLE_PIDS = {_SUMMARY: {4: "tac_gia", 8: "nguoi_luu_cuoi"}, _DOC_SUMMARY: {0x0E: "nguoi_quan_ly", 0x0F: "cong_ty"}}
_VT_LPSTR, _VT_LPWSTR = 0x1E, 0x1F
_REC_FILEPASS, _REC_WRITEACCESS = 0x002F, 0x005C
_ASSOC_AUTHOR, _ASSOC_LAST_REV = 6, 7  # ibstAssocAuthor, ibstAssocLastRevBy
_FC_STTBF_ASSOC = 32  # vị trí cặp fc/lcb của SttbfAssoc trong FibRgFcLcb97


def _zip_field(tag: bytes) -> re.Pattern[bytes]:
    return re.compile(rb"(<" + tag + rb"\b[^>]*>)(.*?)(</" + tag + rb">)", re.S)


def _clean_text(s: str) -> str:
    return s.replace("\x00", "").strip()


def _is_person(value: str) -> bool:
    return _clean_text(value) not in ALLOWED


# ---------- file zip ----------

def _zip_people(data: bytes) -> list[tuple[str, str]]:
    out = []
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        names = set(zf.namelist())
        for part, tags in _ZIP_FIELDS.items():
            if part not in names:
                continue
            xml = zf.read(part)
            for tag in tags:
                for m in _zip_field(tag).finditer(xml):
                    out.append((f"{part} {tag.decode()}", _clean_text(m.group(2).decode("utf8", "replace"))))
    return out


def _zip_strip(data: bytes) -> tuple[bytes, int]:
    out = io.BytesIO()
    n = 0
    with zipfile.ZipFile(io.BytesIO(data)) as zi, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for info in zi.infolist():
            d = zi.read(info.filename)
            for tag in _ZIP_FIELDS.get(info.filename, ()):
                pat = _zip_field(tag)
                n += sum(1 for m in pat.finditer(d) if _is_person(m.group(2).decode("utf8", "replace")))
                d = pat.sub(lambda m: m.group(1) + m.group(3) if _is_person(m.group(2).decode("utf8", "replace")) else m.group(0), d)
            zo.writestr(info, d, compress_type=zipfile.ZIP_DEFLATED)
    return out.getvalue(), n


# ---------- file đời cũ (OLE) ----------

def _props(stream: bytes, pids: dict[int, str]):
    """Các chuỗi cần xem trong phần đầu của bộ thuộc tính: [(tên, vị trí byte chuỗi, độ dài byte, giá trị)]."""
    if len(stream) < 48:
        return []
    (sec,) = struct.unpack_from("<I", stream, 44)
    (count,) = struct.unpack_from("<I", stream, sec + 4)
    codepage = 1252
    entries = [struct.unpack_from("<II", stream, sec + 8 + 8 * i) for i in range(count)]
    for pid, off in entries:
        if pid == 1:
            (codepage,) = struct.unpack_from("<H", stream, sec + off + 4)
    out = []
    for pid, off in entries:
        if pid not in pids:
            continue
        at = sec + off
        (vt,) = struct.unpack_from("<H", stream, at)
        (size,) = struct.unpack_from("<I", stream, at + 4)
        if vt == _VT_LPSTR:
            raw = stream[at + 8: at + 8 + size]
            enc = "utf8" if codepage == 65001 else f"cp{codepage}"
            try:
                text = raw.decode(enc, "replace")
            except LookupError:
                text = raw.decode("latin-1")
            out.append((pids[pid], at + 8, size, _clean_text(text)))
        elif vt == _VT_LPWSTR:
            out.append((pids[pid], at + 8, size * 2, _clean_text(stream[at + 8: at + 8 + size * 2].decode("utf-16-le", "replace"))))
    return out


def _writeaccess(wb: bytes):
    """Bản ghi WRITEACCESS của .xls: (vị trí byte chữ, độ dài, 1 hay 2 byte mỗi chữ, giá trị). None nếu file mã hóa hoặc không có."""
    pos = 0
    while pos + 4 <= len(wb):
        rec, ln = struct.unpack_from("<HH", wb, pos)
        if rec == _REC_FILEPASS:
            return None
        if rec == _REC_WRITEACCESS and ln >= 3:
            (cch,) = struct.unpack_from("<H", wb, pos + 4)
            wide = 2 if wb[pos + 6] & 1 else 1
            start = pos + 7
            raw = wb[start: start + cch * wide]
            return start, cch * wide, wide, _clean_text(raw.decode("utf-16-le" if wide == 2 else "latin-1", "replace"))
        pos += 4 + ln
    return None


def _xls_encrypted(wb: bytes) -> bool:
    """.xls có mật khẩu: có FILEPASS trong phần đầu (trước EOF đầu tiên); WRITEACCESS sau đó bị mã hóa."""
    pos = 0
    while pos + 4 <= len(wb):
        rec, ln = struct.unpack_from("<HH", wb, pos)
        if rec == _REC_FILEPASS:
            return True
        if rec == 0x000A:  # EOF của phần đầu
            return False
        pos += 4 + ln
    return False


def _assoc(word: bytes, table: bytes):
    """SttbfAssoc của .doc: [(tên, vị trí byte chữ trong bảng, độ dài, giá trị)]."""
    (csw,) = struct.unpack_from("<H", word, 32)
    p = 34 + csw * 2
    (cslw,) = struct.unpack_from("<H", word, p)
    p += 2 + cslw * 4
    (cb,) = struct.unpack_from("<H", word, p)
    if cb <= _FC_STTBF_ASSOC:
        return []
    fc, lcb = struct.unpack_from("<II", word, p + 2 + _FC_STTBF_ASSOC * 8)
    if not lcb or fc + 6 > len(table):
        return []
    ext, cdata, cbextra = struct.unpack_from("<HHH", table, fc)
    if ext != 0xFFFF:
        return []
    out = []
    q = fc + 6
    for i in range(cdata):
        (cch,) = struct.unpack_from("<H", table, q)
        if i in (_ASSOC_AUTHOR, _ASSOC_LAST_REV):
            name = "tac_gia" if i == _ASSOC_AUTHOR else "nguoi_sua_cuoi"
            out.append((name, q + 2, cch * 2, _clean_text(table[q + 2: q + 2 + cch * 2].decode("utf-16-le", "replace"))))
        q += 2 + cch * 2 + cbextra
    return out


def _table_name(word: bytes) -> str:
    (flags,) = struct.unpack_from("<H", word, 0x0A)
    return "1Table" if flags & 0x0200 else "0Table"


def _ole_fields(o: olefile.OleFileIO):
    """[(luồng, tên, vị trí, độ dài, giá trị, byte thay)]."""
    out = []
    for stream, pids in _OLE_PIDS.items():
        if o.exists(stream):
            out += [(stream, n, at, ln, v, b"\x00") for n, at, ln, v in _props(o.openstream(stream).read(), pids)]
    if o.exists("Workbook"):
        w = _writeaccess(o.openstream("Workbook").read())
        if w:
            at, ln, wide, v = w
            out.append(("Workbook", "nguoi_luu_writeaccess", at, ln, v, b" " if wide == 1 else b" \x00"))
    if o.exists("WordDocument"):
        word = o.openstream("WordDocument").read()
        table = _table_name(word)
        if o.exists(table):
            out += [(table, n, at, ln, v, b" \x00") for n, at, ln, v in _assoc(word, o.openstream(table).read())]
    return out


def _ole_people(data: bytes) -> list[tuple[str, str]]:
    with olefile.OleFileIO(io.BytesIO(data)) as o:
        return [(f"{stream.lstrip(chr(5))} {name}", v) for stream, name, _, _, v, _ in _ole_fields(o)]


def _ole_strip(data: bytes) -> tuple[bytes, int]:
    with olefile.OleFileIO(io.BytesIO(data)) as o:
        fields = [f for f in _ole_fields(o) if _is_person(f[4])]
        streams = {s: bytearray(o.openstream(s).read()) for s in {f[0] for f in fields}}
    if not fields:
        return data, 0
    for stream, _, at, ln, _, fill in fields:
        streams[stream][at: at + ln] = (fill * ln)[:ln]
    fd, tmp = tempfile.mkstemp(suffix=".ole")
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(data)
        with olefile.OleFileIO(tmp, write_mode=True) as o:
            for stream, d in streams.items():
                o.write_stream(stream, bytes(d))
        with open(tmp, "rb") as f:
            return f.read(), len(fields)
    finally:
        os.unlink(tmp)


# ---------- dùng chung ----------

def kind(data: bytes) -> str:
    """zip, ole, ole_ma_hoa (gói Office có mật khẩu) hoặc khac."""
    if data[:4] == b"PK\x03\x04" and zipfile.is_zipfile(io.BytesIO(data)):
        return "zip"
    if olefile.isOleFile(io.BytesIO(data)):
        with olefile.OleFileIO(io.BytesIO(data)) as o:
            if o.exists("EncryptionInfo") or (o.exists("Workbook") and _xls_encrypted(o.openstream("Workbook").read())):
                return "ole_ma_hoa"
            return "ole"
    return "khac"


def people(data: bytes) -> list[tuple[str, str]]:
    """[(vị trí, giá trị)] cho mọi ô tên người trong thông tin file, kể cả ô rỗng."""
    k = kind(data)
    if k == "zip":
        return _zip_people(data)
    if k in ("ole", "ole_ma_hoa"):  # file mã hóa: chỉ phần nằm ngoài gói mã hóa
        return _ole_people(data)
    return []


def strip(data: bytes) -> tuple[bytes, int]:
    """Gỡ tên người khỏi thông tin file. Trả (byte mới, số ô đã gỡ). File hỏng trả nguyên; file mã hóa chỉ gỡ được phần ngoài gói mã hóa."""
    k = kind(data)
    if k == "zip":
        return _zip_strip(data)
    if k in ("ole", "ole_ma_hoa"):
        return _ole_strip(data)
    return data, 0
