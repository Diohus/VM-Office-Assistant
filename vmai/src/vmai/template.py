"""Mô tả biểu mẫu tổng quát (Q32). Bố cục nằm trong file mô tả, không nằm trong code.

`TemplateSpec` nói biểu mẫu tổ chức thế nào. `FormFacts` là những gì đọc được từ một file cụ thể, dạng JSON.
`read_form` là hàm duy nhất ở đây chạm file, và chỉ đọc.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import json
import re
from collections.abc import Mapping
from pathlib import Path
from typing import Any, Literal

import openpyxl
import yaml
from openpyxl.utils import column_index_from_string
from openpyxl.utils.cell import coordinate_from_string
from pydantic import BaseModel, ConfigDict, Field, ValidationError, model_validator

from .errors import TemplateSpecError
from .phi import is_person_header

_READ_DATE_RE = re.compile(r"\d{1,2}\s*[/\-.]\s*\d{1,2}\s*[/\-.]\s*\d{4}")


class _Model(BaseModel):
    model_config = ConfigDict(extra="forbid", frozen=True)


class KeyColumn(_Model):
    column: str
    dim: str
    kind: Literal["value", "date_sequence"] = "value"
    start_cell: str | None = None  # date_sequence: ô bắt đầu, mỗi dòng thêm một ngày

    @model_validator(mode="after")
    def _needs_start(self) -> KeyColumn:
        if self.kind == "date_sequence" and not self.start_cell:
            raise ValueError("date_sequence cần start_cell")
        return self


class ValueColumn(_Model):
    column: str
    field: str
    dims: dict[str, str] = Field(default_factory=dict)
    type: Literal["number", "text"] = "number"
    # Cột chữ mặc định không được điền. Phải khai báo rõ là danh mục, không phải thông tin về con người.
    category: bool = False

    @model_validator(mode="after")
    def _text_needs_category(self) -> ValueColumn:
        if self.type == "text" and not self.category:
            raise ValueError(f"Cột chữ {self.column} phải khai báo category, hoặc đưa vào manual_columns")
        return self


class TableSpec(_Model):
    sheet: str
    anchor: str
    header_rows: tuple[int, ...]
    first_row: int
    last_row: int
    key_columns: tuple[KeyColumn, ...]
    value_columns: tuple[ValueColumn, ...]
    manual_columns: tuple[str, ...] = ()  # nhân viên tự viết, AI không đề xuất


class LimitsSpec(_Model):
    sheet: str
    first_row: int
    dims: dict[str, str]  # chiều -> cột trong sheet giới hạn
    target: str
    sd: str
    unit: str | None = None
    lot: str | None = None
    dim_value_map: dict[str, dict[str, str]] = Field(default_factory=dict)


class OptionsSpec(_Model):
    applies_to: str
    sheet: str
    column: str
    first_row: int


class SourceField(_Model):
    names: tuple[str, ...]
    aliases: tuple[str, ...] = ()
    required: bool = False


class PeriodSpec(_Model):
    cell: str
    field: str


# Phiên bản định dạng của mô tả đoán ra (Q117). Tăng khi thêm trường làm đổi màu.
# 2: có level_answer_columns, alternative_value_fields (quy tắc màu 0.6).
SPEC_FORMAT = 2
# Mô tả đoán ra do backend lưu lại phải có đủ các trường này, viết rõ ra, kể cả khi rỗng.
# Thiếu thì trần vàng (Q114, Q115) mất mà không ai biết, nên không đọc.
_INFERRED_REQUIRED = ("spec_format", "user_chosen_fields", "level_answer_columns", "alternative_value_fields")


class TemplateSpec(_Model):
    spec_id: str
    spec_format: int = 1
    version: int
    # "inferred": do infer_spec đoán, phần ghép không cao hơn vàng cho tới khi có hồ sơ được xác nhận.
    origin: Literal["declared", "inferred"] = "declared"
    table: TableSpec
    source_fields: dict[str, SourceField]
    # Bảng đồng nghĩa khai báo sẵn (Q16): chiều -> giá trị chuẩn -> các cách ghi.
    dim_values: dict[str, dict[str, tuple[str, ...]]] = Field(default_factory=dict)
    dim_aliases: dict[str, dict[str, tuple[str, ...]]] = Field(default_factory=dict)
    date_fields: tuple[str, ...] = ()
    limits: LimitsSpec | None = None
    options: tuple[OptionsSpec, ...] = ()
    period: PeriodSpec | None = None
    unit_field: str | None = None
    lot_field: str | None = None
    # Cột do người dùng tự chọn ở màn xác nhận (US-1.13). Màu không đổi, chỉ thêm lý do để người xem biết.
    user_chosen_fields: tuple[str, ...] = ()
    # Cột biểu mẫu ghép được nhờ người dùng trả lời "Cùng một thứ" ở LEVEL_NAME_DIFFERS (chữ cột).
    # Q114, Q115: cột ở đây và cột trong user_chosen_fields tối đa vàng, kể cả khi đã có hồ sơ.
    level_answer_columns: tuple[str, ...] = ()
    # Trường kết quả mà bộ đoán chọn giữa nhiều cột số của nguồn, ví dụ "Kết quả" và "KQ máy 1".
    # T-AI-12 (chờ PM ghi số Q): tối đa vàng mãi. Chọn nhầm cột gần giống thì số vẫn trong giới hạn, Q17 không bắt được.
    alternative_value_fields: tuple[str, ...] = ()

    @model_validator(mode="after")
    def _check_references(self) -> TemplateSpec:
        coordinate_from_string(self.table.anchor)
        for column in self.table.value_columns:
            if column.field not in self.source_fields:
                raise ValueError(f"Cột {column.column}: không có trường nguồn {column.field}")
            for dim in column.dims:
                if dim not in self.source_fields:
                    raise ValueError(f"Cột {column.column}: không có trường nguồn cho chiều {dim}")
        for key in self.table.key_columns:
            if key.dim not in self.source_fields:
                raise ValueError(f"Khóa {key.column}: không có trường nguồn cho chiều {key.dim}")
        return self


def load_spec(source: Path | str | Mapping[str, Any]) -> TemplateSpec:
    """Đọc mô tả từ file YAML hoặc từ dict."""
    try:
        data = source if isinstance(source, Mapping) else yaml.safe_load(Path(source).read_text(encoding="utf-8"))
        if isinstance(data, Mapping) and data.get("origin") == "inferred":
            missing = [key for key in _INFERRED_REQUIRED if key not in data]
            if missing or data.get("spec_format") != SPEC_FORMAT:
                # Q117: backend coi như chưa có hồ sơ và cho người dùng xác nhận lại.
                raise TemplateSpecError(
                    "SPEC_OUTDATED",
                    f"Mô tả đoán ra đã cũ hoặc thiếu trường (cần định dạng {SPEC_FORMAT}, thiếu: {', '.join(missing) or 'không'}). Cần xác nhận lại cách ghép.",
                )
        return TemplateSpec.model_validate(data)
    except (ValidationError, ValueError, OSError, yaml.YAMLError) as exc:
        raise TemplateSpecError("SPEC_INVALID", f"Mô tả biểu mẫu không hợp lệ: {exc}") from exc


class Limit(_Model):
    target: float
    sd: float
    unit: str | None = None
    lot: str | None = None

    @property
    def low_3sd(self) -> float:
        return self.target - 3 * self.sd

    @property
    def high_3sd(self) -> float:
        return self.target + 3 * self.sd


class FormColumn(_Model):
    letter: str
    col_path: tuple[str, ...]
    field: str
    dims: dict[str, str]
    type: Literal["number", "text"]
    limit: Limit | None = None
    options: tuple[str, ...] = ()


class ExistingCell(_Model):
    value: str | None = None
    is_number: bool = False
    formula: str | None = None


class FormRow(_Model):
    key: tuple[str, ...]
    rows: tuple[int, ...]  # hơn một dòng là trùng khóa


class FormFacts(_Model):
    spec: TemplateSpec
    file_name: str
    sha256: str
    fingerprint: str
    columns: tuple[FormColumn, ...]
    key_paths: dict[str, tuple[str, ...]]
    manual_columns: dict[str, tuple[str, ...]] = Field(default_factory=dict)  # cột để nhân viên tự viết -> tiêu đề
    rows: tuple[FormRow, ...]
    existing: dict[str, ExistingCell]
    period: dt.date | None = None

    def row_map(self) -> dict[tuple[str, ...], tuple[int, ...]]:
        return {r.key: r.rows for r in self.rows}

    def columns_for(self, record_dims: Mapping[str, str]) -> list[FormColumn]:
        return [c for c in self.columns if all(record_dims.get(k) == v for k, v in c.dims.items())]

    def overlapping(self, cell: FormColumn) -> list[FormColumn]:
        """Các cột số khác có miền ±3SD chồng với cột này."""
        if cell.limit is None:
            return []
        return [
            other
            for other in self.columns
            if other.letter != cell.letter
            and other.limit is not None
            and cell.limit.low_3sd <= other.limit.high_3sd
            and other.limit.low_3sd <= cell.limit.high_3sd
        ]

    def has_independent_check(self, cell: FormColumn, source_unit: str | None = None, relax_units: bool = False) -> bool:
        """Q17, theo từng cột: miền ±3SD không chồng với mọi cột khác, kể cả khác đơn vị.

        `relax_units` chỉ để đo Q30, không dùng trong sản phẩm: nới ra thì có số sai bị xanh khi nguồn ghi sai đơn vị.
        """
        if cell.type == "text":
            return bool(cell.options)
        if cell.limit is None or cell.limit.sd <= 0:
            return False
        overlaps = self.overlapping(cell)
        if not overlaps:
            return True
        if not relax_units or not source_unit or not cell.limit.unit:
            return False
        return all(o.limit is not None and o.limit.unit and o.limit.unit != cell.limit.unit for o in overlaps)


def structure_fingerprint(sheet_names: list[str], spec: TemplateSpec, merged: list[str], headers: list[list[Any]]) -> str:
    """Dấu vân cấu trúc: mô tả, tên sheet, vùng gộp và chữ ở dòng tiêu đề. Không tính số liệu."""
    payload = json.dumps(
        {"spec": spec.spec_id, "sheets": sheet_names, "merged": sorted(merged), "headers": headers}, ensure_ascii=False, default=str
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def sha256_file(path: Path | str) -> str:
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def as_date(value: Any) -> dt.date | None:
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, str):
        for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
            try:
                return dt.datetime.strptime(value.strip(), fmt).date()
            except ValueError:
                continue
    return None


def classify_text(text: str) -> tuple[str, str]:
    """Chữ đọc từ ảnh hoặc PDF: ngày thì đưa về dạng ISO, còn lại giữ nguyên."""
    stripped = text.strip()
    if _READ_DATE_RE.fullmatch(stripped.replace(" ", "")):
        day = as_date(stripped.replace(" ", "").replace("-", "/").replace(".", "/"))
        if day is not None:
            return "date", day.isoformat()
    return "text", stripped


def key_text(value: Any) -> str:
    """Dạng chuẩn của một phần khóa dòng: ngày theo ISO, số nguyên bỏ .0, chữ giữ nguyên đã cắt khoảng trắng."""
    day = as_date(value)
    if day is not None:
        return day.isoformat()
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _col_path(ws: Any, letter: str, header_rows: tuple[int, ...]) -> tuple[str, ...]:
    """Tiêu đề cột theo thứ bậc. Ô gộp lấy chữ ở ô trái trên của vùng gộp."""
    path: list[str] = []
    column = column_index_from_string(letter)
    for row in header_rows:
        value = ws.cell(row=row, column=column).value
        if value is None:
            for merged in ws.merged_cells.ranges:
                if merged.min_row <= row <= merged.max_row and merged.min_col <= column <= merged.max_col:
                    value = ws.cell(row=merged.min_row, column=merged.min_col).value
                    break
        if value not in (None, "") and (not path or path[-1] != str(value)):
            path.append(str(value))
    return tuple(path)


def _read_limits(wb: Any, spec: TemplateSpec) -> dict[tuple[tuple[str, str], ...], Limit]:
    out: dict[tuple[tuple[str, str], ...], Limit] = {}
    raw = spec.limits
    if raw is None or raw.sheet not in wb.sheetnames:
        return out
    ws = wb[raw.sheet]
    for row in range(raw.first_row, ws.max_row + 1):
        target, sd = ws[f"{raw.target}{row}"].value, ws[f"{raw.sd}{row}"].value
        if not isinstance(target, (int, float)) or not isinstance(sd, (int, float)):
            continue
        dims = []
        for dim, letter in raw.dims.items():
            value = key_text(ws[f"{letter}{row}"].value)
            dims.append((dim, raw.dim_value_map.get(dim, {}).get(value, value)))
        unit = ws[f"{raw.unit}{row}"].value if raw.unit else None
        lot = ws[f"{raw.lot}{row}"].value if raw.lot else None
        out[tuple(sorted(dims))] = Limit(target=float(target), sd=float(sd), unit=str(unit) if unit else None, lot=str(lot) if lot else None)
    return out


def _read_rows(ws: Any, spec: TemplateSpec) -> tuple[FormRow, ...]:
    table = spec.table
    rows: dict[tuple[str, ...], list[int]] = {}
    for offset, row in enumerate(range(table.first_row, table.last_row + 1)):
        parts: list[str] | None = []
        for key in table.key_columns:
            if key.kind == "date_sequence":
                # Cột khóa dạng =A8+1 có thể chưa có giá trị đã lưu, nên tính lại từ ô bắt đầu.
                start = as_date(ws[key.start_cell].value)
                if start is None:
                    raise TemplateSpecError("SPEC_MISMATCH", f"Ô {key.start_cell} không phải ngày")
                day = start + dt.timedelta(days=offset)
                if day.month != start.month:
                    parts = None
                    break
                parts.append(day.isoformat())
            else:
                value = ws[f"{key.column}{row}"].value
                if value is None:  # ô gộp dọc, ví dụ một ngày cho hai ca: lấy giá trị ô gốc
                    column = column_index_from_string(key.column)
                    for merged in ws.merged_cells.ranges:
                        if merged.min_row <= row <= merged.max_row and merged.min_col <= column <= merged.max_col:
                            value = ws.cell(row=merged.min_row, column=merged.min_col).value
                            break
                if value in (None, "") or (isinstance(value, str) and value.startswith("=")):
                    parts = None
                    break
                parts.append(key_text(value))
        if parts:
            rows.setdefault(tuple(parts), []).append(row)
    return tuple(FormRow(key=k, rows=tuple(v)) for k, v in rows.items())


def read_form(path: Path | str, spec: TemplateSpec) -> FormFacts:
    """Đọc biểu mẫu theo mô tả. Không mạng, không chạy macro, không theo liên kết ngoài, không ghi file."""
    path = Path(path)
    try:
        wb = openpyxl.load_workbook(path, data_only=False, keep_vba=False, keep_links=False)
    except Exception as exc:
        raise TemplateSpecError("TEMPLATE_UNREADABLE", f"Không mở được biểu mẫu: {type(exc).__name__}") from exc
    table = spec.table
    if table.sheet not in wb.sheetnames:
        raise TemplateSpecError("SPEC_MISMATCH", f"Biểu mẫu không có sheet {table.sheet}")
    ws = wb[table.sheet]
    limits = _read_limits(wb, spec)

    options_by_letter: dict[str, tuple[str, ...]] = {}
    for option in spec.options:
        if option.sheet in wb.sheetnames:
            source = wb[option.sheet]
            values = [source[f"{option.column}{r}"].value for r in range(option.first_row, source.max_row + 1)]
            options_by_letter[option.applies_to] = tuple(str(v) for v in values if v not in (None, ""))

    columns = []
    existing: dict[str, ExistingCell] = {}
    manual = {letter: _col_path(ws, letter, table.header_rows) for letter in table.manual_columns}
    for column in table.value_columns:
        col_path = _col_path(ws, column.column, table.header_rows)
        if not col_path:
            raise TemplateSpecError("SPEC_MISMATCH", f"Cột {column.column} không có tiêu đề ở dòng {table.header_rows}")
        if column.type == "text" and any(is_person_header(part) for part in col_path):
            # Mô tả khai sai cũng không qua được: tiêu đề của file thật là thứ quyết định.
            manual[column.column] = col_path
            continue
        limit = limits.get(tuple(sorted(column.dims.items()))) if column.type == "number" else None
        columns.append(
            FormColumn(
                letter=column.column,
                col_path=col_path,
                field=column.field,
                dims=column.dims,
                type=column.type,
                limit=limit,
                options=options_by_letter.get(column.column, ()),
            )
        )
        for row in range(table.first_row, table.last_row + 1):
            value = ws[f"{column.column}{row}"].value
            if value in (None, ""):
                continue
            if isinstance(value, str) and value.startswith("="):
                existing[f"{column.column}{row}"] = ExistingCell(formula=value)
            else:
                existing[f"{column.column}{row}"] = ExistingCell(value=str(value), is_number=isinstance(value, (int, float)))

    period = None
    if spec.period:
        period = as_date(ws[spec.period.cell].value)
        if period is None:
            raise TemplateSpecError("SPEC_MISMATCH", f"Ô {spec.period.cell} của biểu mẫu không phải ngày")

    headers = [[ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)] for r in table.header_rows]
    merged = [str(m) for m in ws.merged_cells.ranges if m.min_row in table.header_rows]
    return FormFacts(
        spec=spec,
        file_name=path.name,
        sha256=sha256_file(path),
        fingerprint=structure_fingerprint(list(wb.sheetnames), spec, merged, headers),
        columns=tuple(columns),
        key_paths={k.dim: _col_path(ws, k.column, table.header_rows) or (k.dim,) for k in table.key_columns},
        manual_columns=manual,
        rows=_read_rows(ws, spec),
        existing=existing,
        period=period,
    )
