"""Đọc file nguồn Excel bằng code, trả về JSON. Không mạng, không key, không chạy macro, không tính lại công thức."""
from __future__ import annotations

import datetime as dt
import zipfile
from pathlib import Path
from typing import Any, Literal

import openpyxl
from pydantic import BaseModel, ConfigDict

from .errors import SourceFileError
from .template import sha256_file

_OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
_ENCRYPTED_STREAM = "EncryptedPackage".encode("utf-16-le")
# Giới hạn chống file nén cực lớn và bảng quá dài. Dev có cổng đầu vào riêng; đây là lớp thứ hai.
MAX_UNCOMPRESSED_BYTES = 200 * 1024 * 1024
MAX_ROWS = 200_000


class CellRead(BaseModel):
    """Với nguồn ảnh, PDF: ô đọc được ở đâu và lần đọc thứ hai ra gì. `second_text = None` là chưa đọc được."""

    model_config = ConfigDict(frozen=True)

    page: int
    bbox: tuple[float, float, float, float]
    reader: str
    second_reader: str | None = None
    second_text: str | None = None


class SourceCell(BaseModel):
    model_config = ConfigDict(frozen=True)

    a1: str
    kind: Literal["empty", "number", "text", "date", "bool"]
    value: str = ""  # số giữ dạng chữ để JSON không làm đổi giá trị; ngày theo ISO
    formula: str | None = None
    read: CellRead | None = None

    @property
    def is_text(self) -> bool:
        return self.kind == "text"

    @property
    def is_empty(self) -> bool:
        return self.kind == "empty" or (self.kind == "text" and not self.value.strip())

    @property
    def text(self) -> str:
        if self.kind == "date":
            return dt.date.fromisoformat(self.value).strftime("%d/%m/%Y")
        return self.value


class SourceSheet(BaseModel):
    model_config = ConfigDict(frozen=True)

    name: str
    rows: tuple[tuple[SourceCell, ...], ...]
    merged: tuple[str, ...] = ()
    printed_lines: tuple[str, ...] = ()  # dòng chữ ngoài bảng đọc từ ảnh, PDF, ví dụ dòng tổng in trên phiếu


class SourceFacts(BaseModel):
    model_config = ConfigDict(frozen=True)

    file_name: str
    sha256: str
    sheets: tuple[SourceSheet, ...]
    has_macros: bool
    hidden_sheets: int
    kind: Literal["xlsx", "pdf_text", "pdf_image", "image"] = "xlsx"
    low_quality: bool = False
    handwriting_declared: bool = False  # R5b: người dùng khai báo ảnh có chữ viết tay (Q28)
    pdf_text_verified: bool = False  # Q98: chỉ bước so lớp chữ với OCR được đặt cờ này. Hiện chưa có bước đó
    scanned_text_layer: bool = False  # R11: lớp chữ PDF nằm trên ảnh quét, do bên đọc file phát hiện


def _check_container(path: Path) -> None:
    data = path.read_bytes()
    if data[:8] == _OLE_MAGIC:
        # .xlsx có mật khẩu được đóng gói OLE, có luồng EncryptedPackage; .xls cũ thì không.
        if _ENCRYPTED_STREAM in data:
            raise SourceFileError("FILE_PASSWORD", "File có mật khẩu, chưa đọc được")
        raise SourceFileError("FILE_LEGACY_XLS", "File .xls đời cũ, cần chuyển đổi trước")
    if not zipfile.is_zipfile(path):
        raise SourceFileError("FILE_CORRUPT", "File hỏng hoặc không phải Excel")
    try:
        with zipfile.ZipFile(path) as zf:
            if sum(i.file_size for i in zf.infolist()) > MAX_UNCOMPRESSED_BYTES:
                raise SourceFileError("FILE_TOO_LARGE", "File giải nén quá lớn")
    except zipfile.BadZipFile as exc:
        raise SourceFileError("FILE_CORRUPT", "File hỏng hoặc không phải Excel") from exc


def _to_cell(a1: str, value: Any, formula: str | None) -> SourceCell:
    if value is None:
        return SourceCell(a1=a1, kind="empty", formula=formula)
    if isinstance(value, bool):
        return SourceCell(a1=a1, kind="bool", value=str(value), formula=formula)
    if isinstance(value, (int, float)):
        return SourceCell(a1=a1, kind="number", value=str(value), formula=formula)
    if isinstance(value, dt.datetime):
        return SourceCell(a1=a1, kind="date", value=value.date().isoformat(), formula=formula)
    if isinstance(value, dt.date):
        return SourceCell(a1=a1, kind="date", value=value.isoformat(), formula=formula)
    return SourceCell(a1=a1, kind="text", value=str(value), formula=formula)


def read_source(path: Path | str) -> SourceFacts:
    """Đọc mọi sheet đang hiện. Giá trị công thức là giá trị đã lưu, kèm chữ công thức."""
    path = Path(path)
    _check_container(path)
    try:
        values = openpyxl.load_workbook(path, data_only=True, keep_vba=False, keep_links=False)
        formulas = openpyxl.load_workbook(path, data_only=False, keep_vba=False, keep_links=False)
    except Exception as exc:
        raise SourceFileError("FILE_CORRUPT", f"Không mở được file: {type(exc).__name__}") from exc

    sheets = []
    hidden = 0
    for name in values.sheetnames:
        value_ws, formula_ws = values[name], formulas[name]
        if value_ws.sheet_state != "visible":
            hidden += 1
            continue
        if value_ws.max_row > MAX_ROWS:
            raise SourceFileError("FILE_TOO_LARGE", f"Sheet {name} có quá nhiều dòng")
        rows = []
        for value_row, formula_row in zip(value_ws.iter_rows(), formula_ws.iter_rows()):
            cells = []
            for value_cell, formula_cell in zip(value_row, formula_row):
                raw = formula_cell.value
                formula = raw if isinstance(raw, str) and raw.startswith("=") else None
                cells.append(_to_cell(value_cell.coordinate, value_cell.value, formula))
            rows.append(tuple(cells))
        sheets.append(SourceSheet(name=name, rows=tuple(rows), merged=tuple(str(m) for m in value_ws.merged_cells.ranges)))
    return SourceFacts(file_name=path.name, sha256=sha256_file(path), sheets=tuple(sheets), has_macros=path.suffix.lower() == ".xlsm", hidden_sheets=hidden)
