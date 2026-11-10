"""Cổng đầu vào: quyết định file có được đưa cho bộ đọc hay không, trước khi có thư viện xlsx nào chạm vào nó.

Chỉ dựa vào byte của file, không tin đuôi file. Mọi lý do từ chối có mã riêng để app báo đúng loại lỗi.
"""
from __future__ import annotations

import io
import posixpath
import zipfile
from dataclasses import dataclass, field

import olefile

MAX_FILE_BYTES = 25 * 1024 * 1024
MAX_UNCOMPRESSED_BYTES = 200 * 1024 * 1024
MAX_PARTS = 10_000
MAX_CELLS = 200_000
CHUNK = 1024 * 1024

OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
ZIP_MAGICS = (b"PK\x03\x04",)
SHEET_MAIN = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
MACRO_MAIN = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
_DTD_MARKERS = (b"<!DOCTYPE", b"<!ENTITY", "<!DOCTYPE".encode("utf-16-le"), "<!ENTITY".encode("utf-16-le"))


MSG_CORRUPT = "App không đọc được file này, có thể file bị hỏng hoặc chép thiếu. File gốc của bạn không bị đổi. Hãy thử mở bằng Excel, hoặc lấy lại bản khác."
MSG_PASSWORD = "File đang khóa bằng mật khẩu nên app không mở được. File của bạn không bị đổi. Hãy mở bằng Excel, bỏ mật khẩu ở một bản sao rồi tải lại, hoặc hỏi người giữ file."


def _unsupported(ext: str) -> str:
    kind = f"là {ext}" if ext else "không có đuôi rõ ràng"
    return f"App chỉ mở được file Excel .xlsx, .xlsm và .xls, còn file của bạn {kind}. File của bạn không bị đổi. Hãy chọn file khác."


class Rejected(Exception):
    def __init__(self, code: str, message: str, **detail) -> None:
        super().__init__(code)
        self.code = code
        self.message = message
        self.detail = detail


@dataclass
class Accepted:
    has_macro: bool = False
    warnings: list[dict] = field(default_factory=list)


def _extension(filename: str) -> str:
    return posixpath.splitext(filename.replace("\\", "/"))[1].lower()


def _xls_has_password(data: bytes) -> bool:
    """Trong .xls, bản ghi FILEPASS (0x002F) đứng ngay sau BOF nếu file có mật khẩu mở."""
    with olefile.OleFileIO(io.BytesIO(data)) as ole:
        head = ole.openstream("Workbook").read(4096)
    pos = 0
    for _ in range(8):
        if pos + 4 > len(head):
            return False
        kind, size = int.from_bytes(head[pos:pos + 2], "little"), int.from_bytes(head[pos + 2:pos + 4], "little")
        if kind == 0x002F:
            return True
        pos += 4 + size
    return False


def _reject_ole(data: bytes, filename: str) -> None:
    """File OLE2 có thể là xlsx có mật khẩu, .xls, .doc... phải nhìn tên luồng mới phân biệt được."""
    try:
        ole = olefile.OleFileIO(io.BytesIO(data))
    except OSError as exc:
        raise Rejected("FILE_CORRUPT", MSG_CORRUPT) from exc
    try:
        streams = {"/".join(s) for s in ole.listdir(streams=True, storages=False)}
    finally:
        ole.close()
    if {"EncryptedPackage", "EncryptionInfo"} & streams:
        raise Rejected("FILE_PASSWORD", MSG_PASSWORD)
    if "Workbook" in streams and _xls_has_password(data):
        raise Rejected("FILE_PASSWORD", MSG_PASSWORD)
    if {"Workbook", "Book"} & streams:
        raise Rejected("FILE_LEGACY_XLS", "Đây là file Excel đời cũ (.xls), app hiện chưa mở được. File của bạn không bị đổi. Hãy mở bằng Excel, lưu thành .xlsx rồi tải lại.")
    raise Rejected("FILE_UNSUPPORTED", _unsupported(_extension(filename)))


def _check_names(infos: list[zipfile.ZipInfo]) -> None:
    seen = set()
    for info in infos:
        name = info.filename
        if name.startswith("/") or "\\" in name or ".." in name.split("/") or "\x00" in name:
            raise Rejected("FILE_CORRUPT", MSG_CORRUPT, reason="unsafe_part_name")
        key = name.lower()
        if key in seen:
            raise Rejected("FILE_CORRUPT", MSG_CORRUPT, reason="duplicate_part")
        seen.add(key)
        if info.flag_bits & 0x1:
            raise Rejected("FILE_PASSWORD", MSG_PASSWORD)


def _too_large(limit: str, message: str) -> Rejected:
    return Rejected("FILE_TOO_LARGE", message, limit=limit)


def _scan_parts(zf: zipfile.ZipFile, infos: list[zipfile.ZipInfo]) -> None:
    """Giải nén thật và đếm byte: số khai trong zip có thể nói dối. Từ chối XML có DTD, nguồn của XXE và bomb thực thể."""
    total = 0
    for info in infos:
        scan_xml = info.filename.lower().endswith((".xml", ".rels", ".vml"))
        tail = b""
        try:
            with zf.open(info) as fh:
                while chunk := fh.read(CHUNK):
                    total += len(chunk)
                    if total > MAX_UNCOMPRESSED_BYTES:
                        raise _too_large("uncompressed", "Nội dung file khi mở ra vượt giới hạn 200 MB nên app không mở. File của bạn không bị đổi. Hãy tách file hoặc hỏi người quản lý app.")
                    if scan_xml:
                        window = tail + chunk
                        if any(marker in window for marker in _DTD_MARKERS):
                            raise Rejected("FILE_UNSAFE", "File có thành phần app không cho phép nên app không mở. File của bạn không bị đổi. Hãy mở bằng Excel, lưu thành file khác rồi tải lại, hoặc hỏi người quản lý app.", part=info.filename)
                        tail = chunk[-32:]
        except (zipfile.BadZipFile, EOFError, OSError, RuntimeError, NotImplementedError) as exc:
            raise Rejected("FILE_CORRUPT", MSG_CORRUPT) from exc
        except Exception as exc:
            if isinstance(exc, Rejected):
                raise
            raise Rejected("FILE_CORRUPT", MSG_CORRUPT) from exc


def _main_content_type(zf: zipfile.ZipFile) -> str | None:
    from lxml import etree

    try:
        root = etree.fromstring(zf.read("[Content_Types].xml"), etree.XMLParser(resolve_entities=False, no_network=True))
    except (KeyError, etree.XMLSyntaxError) as exc:
        raise Rejected("FILE_CORRUPT", MSG_CORRUPT) from exc
    for el in root:
        if el.get("PartName") == "/xl/workbook.xml":
            return el.get("ContentType")
    return None


def check(data: bytes, filename: str) -> Accepted:
    if len(data) > MAX_FILE_BYTES:
        raise _too_large("file", f"File nặng {len(data) / 1024 / 1024:.0f} MB, vượt giới hạn 25 MB nên app không mở. File của bạn không bị đổi. Hãy tách file hoặc hỏi người quản lý app.")
    ext = _extension(filename)
    if data[:8] == OLE_MAGIC:
        _reject_ole(data, filename)
    if not data.startswith(ZIP_MAGICS):
        if ext in (".xlsx", ".xlsm"):
            raise Rejected("FILE_CORRUPT", MSG_CORRUPT)
        raise Rejected("FILE_UNSUPPORTED", _unsupported(ext))
    try:
        zf = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile as exc:
        if ext in (".xlsx", ".xlsm"):
            raise Rejected("FILE_CORRUPT", MSG_CORRUPT) from exc
        raise Rejected("FILE_UNSUPPORTED", _unsupported(ext)) from exc
    with zf:
        infos = zf.infolist()
        if len(infos) > MAX_PARTS:
            raise _too_large("parts", "File có quá nhiều thành phần bên trong nên app không mở. File của bạn không bị đổi. Hãy tách file hoặc hỏi người quản lý app.")
        _check_names(infos)
        names = {i.filename for i in infos}
        if "[Content_Types].xml" not in names or "xl/workbook.xml" not in names:
            if ext in (".xlsx", ".xlsm"):
                raise Rejected("FILE_CORRUPT", MSG_CORRUPT)
            raise Rejected("FILE_UNSUPPORTED", _unsupported(ext))
        _scan_parts(zf, infos)
        content_type = _main_content_type(zf)
        if content_type not in (SHEET_MAIN, MACRO_MAIN):
            raise Rejected("FILE_UNSUPPORTED", _unsupported(ext))
        has_macro = "xl/vbaProject.bin" in names
        accepted = Accepted(has_macro=has_macro)
        if has_macro or content_type == MACRO_MAIN:
            accepted.warnings.append({
                "code": "FILE_MACRO_IGNORED",
                "message": "File này có phần tự động hóa (macro). App không chạy phần đó, nhưng vẫn xem được dữ liệu và công thức. File của bạn không bị đổi." if has_macro
                else "File này có thể có phần tự động hóa (macro). App không chạy phần đó, nhưng vẫn xem được dữ liệu và công thức. File của bạn không bị đổi.",
            })
        return accepted
