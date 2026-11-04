"""Đoán mô tả biểu mẫu lạ (Q53). Thuần, không I/O, không LLM: nhận `SourceFacts` của biểu mẫu và của nguồn.

Kết quả chỉ là đề xuất cho màn xác nhận hồ sơ. Mô tả đoán ra mang `origin = "inferred"`, nên phần ghép
không bao giờ cao hơn vàng cho tới khi có người xác nhận. Không chắc thì hỏi, không đoán.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import re
from collections import Counter
from dataclasses import dataclass, field

from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.cell import coordinate_from_string, range_boundaries
from pydantic import BaseModel, ConfigDict

from .errors import FeatureNotReleased
from .mapping import normalize_text
from .numbers import ParseStatus, parse_number
from .phi import is_person_header, phi_kind
from .profile import TemplateProfile
from .source_xlsx import SourceCell, SourceFacts, SourceSheet
from .template import TemplateSpec, as_date, key_text, load_spec

MIN_DATA_ROWS = 2
MIN_FILL = 0.7
KEY_OVERLAP = 0.5
MAX_CATEGORY_VALUES = 6
EXTRA_DIMENSION_SHARE = 0.2
_ORDINAL_HEADERS = ("stt", "tt", "so tt", "so thu tu")
_STOP_WORDS = ("tong", "cong", "trung binh", "sd", "cv", "so lan", "nguoi lap", "nguoi kiem", "ky ten", "lech so", "total")
_NOTE_HEADERS = ("ghi chu", "nhan xet", "note", "dien giai", "ly do")
_LEVEL = re.compile(r"\b(?:muc|level|lv|l|m)\s*(\d+)\b")
_NUMBER_NAMES = {
    "ket qua": ("kq", "result", "gia tri", "tri so"),
    "so luong": ("sl", "so luong xuat", "qty", "quantity"),
    "don gia": ("gia", "gia tien", "price"),
}
_UNIT_NAMES = ("don vi", "dvt", "unit")
_LOT_NAMES = ("lo", "so lo", "lot")
_TARGET_NAMES = ("dich", "gia tri dich", "target", "mean")
_SD_NAMES = ("sd", "do lech chuan")


class Question(BaseModel):
    model_config = ConfigDict(frozen=True)

    code: str
    detail: str


class GuessedColumn(BaseModel):
    model_config = ConfigDict(frozen=True)

    letter: str
    col_path: tuple[str, ...]
    source_header: str
    dims: dict[str, str] = {}


class SpecGuess(BaseModel):
    """`status = "ok"` vẫn phải qua màn xác nhận. `"ask"`: có câu hỏi, `spec` có thể rỗng hoặc chỉ gồm phần chắc."""

    model_config = ConfigDict(frozen=True)

    status: str
    spec: TemplateSpec | None = None
    questions: tuple[Question, ...] = ()
    columns: tuple[GuessedColumn, ...] = ()
    computed_columns: tuple[str, ...] = ()
    label_columns: tuple[str, ...] = ()
    manual_columns: tuple[str, ...] = ()


class _Grid:
    def __init__(self, sheet: SourceSheet):
        self.name = sheet.name
        self.cells: dict[tuple[int, int], SourceCell] = {}
        for row in sheet.rows:
            for cell in row:
                column, number = coordinate_from_string(cell.a1)
                self.cells[(number, column_index_from_string(column))] = cell
        self.max_row = max((r for r, _ in self.cells), default=0)
        self.max_col = max((c for _, c in self.cells), default=0)
        self.merge_home: dict[tuple[int, int], tuple[int, int]] = {}
        self.horizontal: set[tuple[int, int]] = set()
        for text in sheet.merged:
            min_col, min_row, max_col, max_row = range_boundaries(text)
            for r in range(min_row, max_row + 1):
                for c in range(min_col, max_col + 1):
                    self.merge_home[(r, c)] = (min_row, min_col)
            if max_col > min_col:
                for c in range(min_col, max_col + 1):
                    self.horizontal.add((min_row, c))

    def own(self, r: int, c: int) -> SourceCell | None:
        cell = self.cells.get((r, c))
        return cell if cell is not None and _has_content(cell) else None

    def get(self, r: int, c: int) -> SourceCell | None:
        """Theo ô gộp về ô gốc."""
        home = self.merge_home.get((r, c))
        return self.own(*home) if home else self.own(r, c)


def _has_content(cell: SourceCell) -> bool:
    return not cell.is_empty or cell.formula is not None


def _is_label(cell: SourceCell) -> bool:
    return cell.is_text and cell.formula is None


@dataclass
class _Table:
    grid: _Grid
    header_rows: tuple[int, ...]
    left: int
    right: int
    data_rows: list[int]
    paths: dict[int, tuple[str, ...]] = field(default_factory=dict)

    @property
    def first_row(self) -> int:
        return self.data_rows[0]

    @property
    def last_row(self) -> int:
        return self.data_rows[-1]

    @property
    def anchor(self) -> str:
        return f"{get_column_letter(self.left)}{self.header_rows[0]}"


def _header_span(grid: _Grid, r: int) -> tuple[int, int] | None:
    """Cột trái, phải của dòng nếu trông giống tiêu đề bảng: chỉ có chữ hoặc ngày, không phải cặp nhãn và giá trị."""
    present = [c for c in range(1, grid.max_col + 1) if grid.get(r, c) is not None]
    own = [c for c in present if grid.own(r, c) is not None]
    if len(own) < 2:
        return None
    for c in own:
        cell = grid.own(r, c)
        if cell.formula is not None or not (_is_label(cell) or cell.kind == "date"):
            return None
        if cell.is_text and cell.value.strip().endswith(":"):
            return None
    left, right = min(present), max(present)
    if len(present) / (right - left + 1) < MIN_FILL:
        return None
    return left, right


def _find_tables(grid: _Grid) -> list[_Table]:
    tables: list[_Table] = []
    r = 1
    while r <= grid.max_row:
        span = _header_span(grid, r)
        if span is None:
            r += 1
            continue
        left, right = span
        header_rows = [r]
        # Dòng dưới là tiêu đề con khi dòng trên có ô gộp ngang trùm lên nó.
        below = r + 1
        while below <= grid.max_row:
            cells = [grid.own(below, c) for c in range(left, right + 1)]
            filled = [c for c in cells if c is not None]
            top_merged = any((below - 1, c) in grid.horizontal for c in range(left, right + 1))
            if filled and all(_is_label(c) for c in filled) and top_merged:
                header_rows.append(below)
                below += 1
            else:
                break
        table = _collect_data(grid, tuple(header_rows), left, right)
        if table is not None:
            tables.append(table)
            r = table.last_row + 1
        else:
            r = below
    return tables


def _collect_data(grid: _Grid, header_rows: tuple[int, ...], left: int, right: int) -> _Table | None:
    rows: list[int] = []
    r = header_rows[-1] + 1
    while r <= grid.max_row:
        cells = [grid.get(r, c) for c in range(left, right + 1)]
        if all(c is None for c in cells):
            break
        first = next((c for c in cells if c is not None), None)
        if first is not None and first.is_text:
            norm = normalize_text(first.value)
            if any(norm.startswith(word) for word in _STOP_WORDS):
                break
        if _header_span(grid, r) is not None and rows and all((c is None or _is_label(c)) for c in cells) and sum(c is not None for c in cells) == right - left + 1:
            break  # gặp tiêu đề của bảng khác
        merged_across = all((r, c) in grid.merge_home for c in range(left, right + 1))
        if not merged_across:  # dòng tên nhóm gộp ngang không phải dữ liệu
            rows.append(r)
        r += 1
    if len(rows) < MIN_DATA_ROWS:
        return None
    table = _Table(grid, header_rows, left, right, rows)
    for c in range(left, right + 1):
        path: list[str] = []
        for row in header_rows:
            cell = grid.get(row, c)
            if cell is None:
                continue
            text = cell.value.strip() if cell.is_text else cell.value
            if text and (not path or path[-1] != text):
                path.append(text)
        table.paths[c] = tuple(path)
    return table


def _strip_units(text: str) -> str:
    return re.sub(r"\(.*?\)", " ", text)


def _value_match(path: tuple[str, ...], value: str) -> int:
    """Mức khớp giữa tiêu đề cột và một giá trị trong nguồn. 0 là không khớp."""
    joined = normalize_text(" ".join(_strip_units(p) for p in path))
    norm = normalize_text(value)
    if not norm or not joined:
        return 0
    tokens = joined.split()
    if norm == joined:
        return 4
    if norm in tokens:
        return 3
    level_h, level_v = set(_LEVEL.findall(joined)), set(_LEVEL.findall(norm))
    if level_h and (level_v or norm.isdigit()) and (level_v or {norm}) <= level_h:
        return 2
    if len(norm) >= 3 and any(t.startswith(norm) for t in tokens):
        return 2
    return 0


def _name_in(header: str, names: tuple[str, ...]) -> bool:
    return normalize_text(header) in names


def _same_number_name(a: str, b: str) -> bool:
    """Hai tiêu đề cùng là một đại lượng: giống nhau, hoặc cùng nhóm trong bảng tên thông dụng."""
    na, nb = normalize_text(a), normalize_text(b)
    if na == nb:
        return True
    return any({na, nb} <= {canon, *aliases} for canon, aliases in _NUMBER_NAMES.items())


@dataclass
class _SourceColumn:
    index: int
    header: str
    values: list[SourceCell]

    @property
    def kinds(self) -> Counter:
        return Counter(c.kind for c in self.values if not c.is_empty)

    @property
    def is_date(self) -> bool:
        k = self.kinds
        return bool(k) and (k["date"] / sum(k.values()) >= 0.8 or sum(as_date(c.value) is not None for c in self.values if c.is_text) / sum(k.values()) >= 0.8)

    @property
    def is_number(self) -> bool:
        """Tính cả số lưu dạng chữ, kể cả dạng nhập nhằng: việc xếp màu là của bước sau."""
        k = self.kinds
        if not k:
            return False
        numeric = k["number"] + sum(parse_number(c.value).status is not ParseStatus.INVALID for c in self.values if c.is_text)
        return numeric / sum(k.values()) >= 0.8

    @property
    def is_small_category(self) -> bool:
        """Số nguyên ít giá trị, như mức 1, 2: có thể là chiều chứ không phải kết quả."""
        numbers = [c.value for c in self.values if c.kind == "number"]
        return bool(numbers) and len(set(numbers)) <= MAX_CATEGORY_VALUES and all(float(v).is_integer() for v in numbers)

    def key_values(self) -> set[str]:
        return {key_text(c.value) for c in self.values if not c.is_empty}

    def distinct_text(self) -> list[str]:
        return sorted({key_text(float(c.value)) if c.kind == "number" else c.value.strip() for c in self.values if c.kind == "number" or (c.is_text and c.value.strip())})


def _source_columns(source: SourceFacts) -> tuple[list[_SourceColumn], _Table | None, SourceSheet | None]:
    best: tuple[_Table, SourceSheet] | None = None
    for sheet in source.sheets:
        for table in _find_tables(_Grid(sheet)):
            if best is None or len(table.data_rows) > len(best[0].data_rows):
                best = (table, sheet)
    if best is None:
        return [], None, None
    table, sheet = best
    columns = []
    for c in range(table.left, table.right + 1):
        path = table.paths.get(c) or ()
        header = str(path[-1]) if path else ""
        if not header or phi_kind(header) or is_person_header(header):
            continue
        values = [cell for r in table.data_rows if (cell := table.grid.own(r, c)) is not None]
        columns.append(_SourceColumn(c - 1, header, values))
    return columns, table, sheet


def _template_keys(table: _Table, column: int) -> tuple[list[str], str | None]:
    """Giá trị khóa của cột. Trả về cả ô bắt đầu nếu cột là dãy ngày dạng công thức."""
    grid = table.grid
    first = grid.own(table.first_row, column)
    if first is not None and first.formula and len(table.data_rows) >= 2:
        match = re.fullmatch(r"=\$?([A-Z]+)\$?(\d+)", first.formula)
        second = grid.own(table.data_rows[1], column)
        if match and second is not None and second.formula and second.formula.endswith("+1"):
            start = f"{match.group(1)}{match.group(2)}"
            cell = grid.cells.get((int(match.group(2)), column_index_from_string(match.group(1))))
            day = as_date(cell.value) if cell is not None and not cell.is_empty else None
            if day is not None:
                out = []
                for offset in range(len(table.data_rows)):
                    current = day + dt.timedelta(days=offset)
                    if current.month == day.month:
                        out.append(current.isoformat())
                return out, start
    values = []
    for r in table.data_rows:
        cell = grid.get(r, column)
        if cell is not None and cell.formula is None:
            values.append(key_text(cell.value))
    return values, None


def _classify_columns(table: _Table):
    """Chia cột theo nội dung: có sẵn đủ (ứng viên khóa, nhãn), công thức, người, ghi chú, còn lại là cột để điền."""
    filled, computed, manual, notes, empty_or_number = [], [], [], [], []
    total = len(table.data_rows)
    for c in range(table.left, table.right + 1):
        path = table.paths.get(c) or ()
        joined = " ".join(str(p) for p in path)
        cells = [table.grid.get(r, c) for r in table.data_rows]
        own = [cell for cell in cells if cell is not None]
        if is_person_header(joined) or phi_kind(joined):
            manual.append(c)
        elif normalize_text(joined) in _NOTE_HEADERS:
            notes.append(c)
        elif sum(cell.formula is not None for cell in own) / total >= 0.5 and not _template_keys(table, c)[1]:
            computed.append(c)
        elif len(own) / total >= 0.9 and not all(cell.kind == "number" and cell.formula is None for cell in own):
            filled.append(c)
        elif len(own) / total >= 0.9:
            filled.append(c)  # toàn số có sẵn: có thể là STT hoặc mã số, xét tiếp ở bước ghép khóa
        else:
            empty_or_number.append(c)
    return filled, computed, manual, notes, empty_or_number


def _is_ordinal(table: _Table, column: int, keys: list[str]) -> bool:
    """Cột số thứ tự trông như khóa nhưng không phải: có thể tình cờ trùng một cột số của nguồn."""
    header = normalize_text(" ".join(str(p) for p in table.paths.get(column, ())))
    return header in _ORDINAL_HEADERS or keys == [str(i + 1) for i in range(len(keys))]


def _pick_keys(table: _Table, filled: list[int], source_cols: list[_SourceColumn]):
    """Chọn cột khóa bằng dữ liệu: giá trị phải trùng với một cột của nguồn. STT không trùng thì không phải khóa."""
    candidates = []
    ambiguous: list[tuple[str, ...]] = []
    for c in filled:
        keys, start = _template_keys(table, c)
        if not keys or _is_ordinal(table, c, keys):
            continue
        best_col, best_score = None, 0.0
        strong = []
        for col in source_cols:
            values = col.key_values()
            if not values:
                continue
            if start or all(as_date(k) for k in keys):
                if not col.is_date:
                    continue
                # Dãy ngày có thể khác tháng với nguồn (sai kỳ là việc của thông báo), nên chỉ cần cột là ngày.
                score = 1.0 if start else len(set(keys) & values) / max(len(set(keys)), 1)
            else:
                score = len(set(keys) & values) / max(len(set(keys)), 1)
            if score >= KEY_OVERLAP:
                strong.append(col.header)
            if score > best_score:
                best_col, best_score = col, score
        if best_col is not None and best_score >= KEY_OVERLAP:
            candidates.append((c, best_col, start, keys))
            if len(strong) > 1:
                ambiguous.append(tuple(strong))
    chosen = []
    for candidate in candidates:
        chosen.append(candidate)
        tuples = list(zip(*(k[3] for k in chosen)))
        if len(set(tuples)) == len(tuples):
            return chosen, True, ambiguous
    return chosen, False, ambiguous


def _limits_spec(template_grids: list[_Grid], data_grid: _Grid, dim_values: dict[str, list[str]]) -> dict | None:
    for grid in template_grids:
        if grid is data_grid:
            continue
        for table in _find_tables(grid):
            headers = {c: normalize_text(" ".join(str(p) for p in table.paths.get(c, ()))) for c in range(table.left, table.right + 1)}
            target = next((c for c, h in headers.items() if h in _TARGET_NAMES), None)
            sd = next((c for c, h in headers.items() if h in _SD_NAMES), None)
            if target is None or sd is None:
                continue
            dims: dict[str, str] = {}
            value_map: dict[str, dict[str, str]] = {}
            for dim, values in dim_values.items():
                for c in range(table.left, table.right + 1):
                    if c in (target, sd):
                        continue
                    raw = {key_text(cell.value) for r in table.data_rows if (cell := table.grid.own(r, c)) is not None and not cell.is_empty}
                    if not raw:
                        continue
                    mapping = {}
                    for item in raw:
                        hits = [v for v in values if v == item or _value_match((item,), v) or (item.isdigit() and _value_match((f"muc {item}",), v))]
                        if len(hits) == 1:
                            mapping[item] = hits[0]
                    if mapping and len(mapping) == len(raw):
                        dims[dim] = get_column_letter(c)
                        value_map[dim] = {k: v for k, v in mapping.items() if k != v}
                        break
            if len(dims) != len(dim_values) or not dims:
                continue
            spec = {"sheet": grid.name, "first_row": table.first_row, "dims": dims, "target": get_column_letter(target), "sd": get_column_letter(sd), "dim_value_map": value_map}
            unit = next((c for c, h in headers.items() if h in _UNIT_NAMES), None)
            lot = next((c for c, h in headers.items() if h in _LOT_NAMES), None)
            if unit:
                spec["unit"] = get_column_letter(unit)
            if lot:
                spec["lot"] = get_column_letter(lot)
            return spec
    return None


def _weak_level_match(path: tuple[str, ...], value: str) -> bool:
    """`M1` với `L1`: cùng số nhưng khác chữ. M1 có thể là máy 1, cấu trúc file không nói được, nên phải hỏi."""
    joined = normalize_text(" ".join(_strip_units(p) for p in path))
    head = re.search(r"\b([a-z]+)\s*(\d+)\b", joined)
    tail = re.search(r"\b([a-z]+)\s*(\d+)\b", normalize_text(value))
    if not head or not tail or head.group(2) != tail.group(2):
        return False
    full = {"muc", "level", "lv"}
    return head.group(1) != tail.group(1) and head.group(1) not in full and tail.group(1) not in full


def infer_spec(template: SourceFacts, source: SourceFacts, accepted: frozenset[str] = frozenset(), *, experimental_image_sources: bool = False) -> SpecGuess:
    """Đoán cách ghép giữa một biểu mẫu lạ và một nguồn. Chỉ điền cột số, cột chữ để người dùng tự viết.

    `accepted`: mã câu hỏi người dùng đã đồng ý với gợi ý, để chạy lại mà không hỏi nữa.
    Còn câu hỏi thì `spec` là None: không có đề xuất nào từ một đoán chưa chắc.
    """
    if source.kind != "xlsx" and not experimental_image_sources:
        raise FeatureNotReleased("Đoán biểu mẫu từ ảnh và PDF")  # Q81
    questions: list[Question] = []

    def ask(code: str, detail: str) -> None:
        if code not in accepted:
            questions.append(Question(code=code, detail=detail))


    source_cols, source_table, _ = _source_columns(source)
    if not source_cols or source_table is None:
        return SpecGuess(status="ask", questions=(Question(code="SOURCE_TABLE_NOT_FOUND", detail="Không thấy bảng số liệu trong file nguồn."),))

    grids = [_Grid(s) for s in template.sheets]
    scored: list[tuple[float, _Table, tuple]] = []
    for grid in grids:
        for table in _find_tables(grid):
            filled, computed, manual, notes, fillable = _classify_columns(table)
            keys, unique, ambiguous = _pick_keys(table, filled, source_cols)
            score = (1.0 if keys else 0.0) * (2 if unique else 1) * len(table.data_rows)
            scored.append((score, table, (filled, computed, manual, notes, fillable, keys, unique, ambiguous)))
    if not scored:
        return SpecGuess(status="ask", questions=(Question(code="TABLE_NOT_FOUND", detail="Không thấy bảng trong biểu mẫu. Biểu mẫu dạng nhãn và ô trống chưa hỗ trợ."),))
    scored.sort(key=lambda item: -item[0])
    _, table, info = scored[0]
    filled, computed, manual, notes, fillable, keys, unique, ambiguous_keys = info
    if len(scored) > 1 and scored[1][0] == scored[0][0] and scored[0][0] > 0:
        ask("TABLE_AMBIGUOUS", "Có nhiều bảng khớp như nhau trong biểu mẫu. Bảng nào cần điền?")
    if not keys:
        return SpecGuess(status="ask", questions=(*questions, Question(code="KEY_NOT_FOUND", detail="Không thấy cột nào của biểu mẫu trùng với cột khóa của nguồn.")))
    if not unique:
        ask("KEY_NOT_UNIQUE", "Khóa dòng bị trùng trong biểu mẫu: không biết số vào dòng nào.")
    for names in ambiguous_keys:
        ask("KEY_SOURCE_AMBIGUOUS", f"Nhiều cột của nguồn có thể là khóa: {', '.join(names)}. Dùng cột nào?")

    key_source = {id(k[1]) for k in keys}
    remaining = [col for col in source_cols if id(col) not in key_source]
    number_cols = [col for col in remaining if col.is_number]
    text_cols = [col for col in remaining if (not col.is_number or col.is_small_category) and not col.is_date and col.distinct_text()]
    date_dims = [col for col in remaining if col.is_date]
    key_columns = {k[0] for k in keys}
    labels = [c for c in filled if c not in key_columns]

    # Cột để điền: cột trống. Cột có sẵn số mà không phải khóa cũng có thể là cột để ghi đè, nhưng phải khớp tên nguồn.
    # Cột trống không có tiêu đề chỉ là cột ngăn.
    targets = [c for c in fillable if table.paths.get(c)]
    for c in list(labels):
        if all((cell := table.grid.get(r, c)) is None or (cell.kind == "number" and cell.formula is None) for r in table.data_rows):
            if any(_same_number_name(str(p), col.header) for p in table.paths.get(c, ()) for col in number_cols):
                targets.append(c)
                labels.remove(c)
    targets.sort()

    paths = {c: table.paths.get(c, ()) for c in targets}
    path_counts = Counter(paths.values())
    duplicated = [c for c in targets if path_counts[paths[c]] > 1]
    if duplicated:
        ask("COLUMN_NAME_REPEATED", f"Nhiều cột cùng tên: {', '.join(get_column_letter(c) for c in duplicated)}.")

    # Chiều: cột chữ của nguồn có giá trị khớp tiêu đề biểu mẫu. Chỉ giữ chiều khớp được ít nhất nửa số cột.
    dim_cols = []
    for col in text_cols:
        hits = sum(any(_value_match(paths[c], v) for v in col.distinct_text()) for c in targets)
        if targets and hits / len(targets) >= 0.5:
            dim_cols.append(col)
    number_cols = [col for col in number_cols if not any(col is d for d in dim_cols)]
    date_header_cols = [col for col in date_dims if all(as_date(p) for c in targets for p in paths[c][-1:])]

    # Nguồn dạng ngang: mỗi cột số của nguồn là một cột của biểu mẫu, ví dụ `ABC M1` với `Abcxyz` / `Mức 1`.
    # Viết cho ảnh và PDF, chưa đo với Excel, nên đi cùng cờ (Q81).
    wide_match = _match_wide_columns(paths, targets, number_cols) if experimental_image_sources and not dim_cols and not date_header_cols else {}

    fields: dict[str, dict] = {}
    field_name: dict[int, str] = {}

    def field_of(col: _SourceColumn, required: bool) -> str:
        if id(col) not in field_name:
            name = f"f{col.index}"
            field_name[id(col)] = name
            fields[name] = {"names": [col.header], "required": required}
        elif required:
            fields[field_name[id(col)]]["required"] = True
        return field_name[id(col)]

    value_columns: list[dict] = []
    guessed: list[GuessedColumn] = []
    weak: set[str] = set()
    unmapped: list[int] = []
    for c in targets:
        path = paths[c]
        if c in duplicated:
            continue
        direct = [col for col in number_cols if any(_same_number_name(str(p), col.header) for p in path)]
        if c in wide_match and len(direct) != 1:
            col = wide_match[c]
            if _weak_level_match(path, col.header):
                weak.add(f"{' / '.join(str(p) for p in path)} là {col.header}")
            value_columns.append({"column": get_column_letter(c), "field": field_of(col, True)})
            guessed.append(GuessedColumn(letter=get_column_letter(c), col_path=tuple(str(p) for p in path), source_header=col.header))
            continue
        if len(direct) == 1 and not dim_cols and not date_header_cols:
            value_columns.append({"column": get_column_letter(c), "field": field_of(direct[0], True)})
            guessed.append(GuessedColumn(letter=get_column_letter(c), col_path=tuple(str(p) for p in path), source_header=direct[0].header))
            continue
        dims: dict[str, str] = {}
        ok = bool(dim_cols or date_header_cols)
        for col in dim_cols:
            scored_values = sorted(((_value_match(path, v), v) for v in col.distinct_text()), key=lambda item: -item[0])
            if not scored_values or scored_values[0][0] == 0 or (len(scored_values) > 1 and scored_values[0][0] == scored_values[1][0]):
                ok = False
                break
            dims[field_of(col, True)] = scored_values[0][1]
            if _weak_level_match(path, scored_values[0][1]):
                weak.add(f"{' / '.join(str(p) for p in path)} là {scored_values[0][1]}")
        if ok:
            for col in date_header_cols:
                day = as_date(path[-1])
                if day is None:
                    ok = False
                    break
                dims[field_of(col, True)] = day.isoformat()
        value_col = _value_column(number_cols)
        if ok and value_col is not None:
            value_columns.append({"column": get_column_letter(c), "field": field_of(value_col, True), "dims": dims})
            guessed.append(GuessedColumn(letter=get_column_letter(c), col_path=tuple(str(p) for p in path), source_header=value_col.header, dims=dims))
        else:
            unmapped.append(c)

    if weak:
        ask("LEVEL_NAME_DIFFERS", f"Cách ghi khác nhau, có đúng không: {'; '.join(sorted(weak))}?")
    # Nguồn còn một chiều mà biểu mẫu không có: mỗi ô đích nhận nhiều số, không biết lấy số nào.
    if value_columns and any(v.get("dims") for v in value_columns):
        used = [k[1] for k in keys] + [c for c in dim_cols] + date_header_cols
        combos_seen = Counter(
            tuple(key_text(cell.value) if (cell := source_table.grid.own(r, col.index + 1)) is not None else "" for col in used) for r in source_table.data_rows
        )
        repeated = sum(n for n in combos_seen.values() if n > 1)
        if repeated / max(len(source_table.data_rows), 1) >= EXTRA_DIMENSION_SHARE:
            ask("SOURCE_EXTRA_DIMENSION", "Nguồn có nhiều số cho cùng một ô của biểu mẫu. Nguồn còn một cách phân loại mà biểu mẫu chưa có.")
    # Hai cột đích cùng một tổ hợp chiều là đoán sai chắc chắn.
    combos = Counter((v["field"], tuple(sorted(v.get("dims", {}).items()))) for v in value_columns)
    if any(count > 1 for count in combos.values()):
        ask("COLUMN_MAPPING_CONFLICT", "Nhiều cột của biểu mẫu ghép vào cùng một loại số.")
    if unmapped:
        ask("COLUMN_NOT_MAPPED", f"Chưa ghép được các cột: {', '.join(' / '.join(str(p) for p in paths[c]) or get_column_letter(c) for c in unmapped)}.")
    if not value_columns:
        ask("NO_FILLABLE_COLUMN", "Không có cột nào của biểu mẫu ghép được với nguồn.")

    letters = lambda items: tuple(get_column_letter(c) for c in items)  # noqa: E731
    base = SpecGuess(status="ask", questions=tuple(questions), columns=tuple(guessed), computed_columns=letters(computed), label_columns=letters(labels), manual_columns=letters([*manual, *notes]))
    if not value_columns:
        return base

    key_specs = []
    date_fields = []
    period = None
    for c, col, start, _ in keys:
        name = field_of(col, True)
        spec_key = {"column": get_column_letter(c), "dim": name}
        if start:
            spec_key.update(kind="date_sequence", start_cell=start)
            period = {"cell": start, "field": name}
        key_specs.append(spec_key)
        if col.is_date:
            date_fields.append(name)
    for col in date_header_cols:
        date_fields.append(field_of(col, True))

    unit_col = next((col for col in source_cols if _name_in(col.header, _UNIT_NAMES)), None)
    lot_col = next((col for col in source_cols if _name_in(col.header, _LOT_NAMES)), None)
    dim_values: dict[str, dict[str, list[str]]] = {}
    for col in dim_cols:
        dim_values[field_of(col, True)] = {v: [v] for v in col.distinct_text()}

    limit_dims = {name: list(values) for name, values in dim_values.items()}
    limits = _limits_spec(grids, table.grid, limit_dims) if limit_dims else None

    header_text = "|".join(" / ".join(str(p) for p in table.paths.get(c, ())) for c in range(table.left, table.right + 1))
    digest = hashlib.sha256(f"{table.grid.name}|{table.anchor}|{header_text}".encode()).hexdigest()[:12]
    data = {
        "spec_id": f"inferred-{digest}",
        "version": 1,
        "origin": "inferred",
        "table": {
            "sheet": table.grid.name,
            "anchor": table.anchor,
            "header_rows": list(table.header_rows),
            "first_row": table.first_row,
            "last_row": table.last_row,
            "key_columns": key_specs,
            "value_columns": value_columns,
            "manual_columns": list(letters([*manual, *notes])),
        },
        "source_fields": fields,
        "dim_values": dim_values,
        "date_fields": sorted(set(date_fields)),
    }
    if period:
        data["period"] = period
    if limits:
        data["limits"] = limits
    if unit_col is not None:
        data["source_fields"][field_of(unit_col, False)] = fields[field_of(unit_col, False)]
        data["unit_field"] = field_of(unit_col, False)
    if lot_col is not None:
        data["lot_field"] = field_of(lot_col, False)
    spec = load_spec(data)
    if questions:
        return base.model_copy(update={"questions": tuple(questions)})
    return base.model_copy(update={"status": "ok", "spec": spec, "questions": ()})


def _match_wide_columns(paths: dict[int, tuple], targets: list[int], number_cols: list[_SourceColumn]) -> dict[int, _SourceColumn]:
    """Ghép một một giữa cột biểu mẫu và cột số của nguồn. Mọi từ của tiêu đề nguồn phải khớp tiêu đề biểu mẫu."""
    out: dict[int, _SourceColumn] = {}
    used: Counter = Counter()
    for c in targets:
        hits = []
        for col in number_cols:
            tokens = normalize_text(col.header).split()
            # Tiêu đề OCR có thể dính chữ: `ABCM1` tách thành `ABC` và `M1`.
            tokens = [t for word in tokens for t in (re.findall(r"[a-z]+|\d+", word) if re.fullmatch(r"[a-z]{3,}[a-z]\d+", word) is None else [word[:-2], word[-2:]])]
            if tokens and all(_value_match(paths[c], t) for t in tokens):
                hits.append(col)
        if len(hits) == 1:
            out[c] = hits[0]
            used[id(hits[0])] += 1
    return {c: col for c, col in out.items() if used[id(col)] == 1}


def _value_column(number_cols: list[_SourceColumn]) -> _SourceColumn | None:
    """Cột số của nguồn dạng dọc. Nhiều cột thì phải có đúng một cột mang tên của kết quả."""
    if len(number_cols) == 1:
        return number_cols[0]
    named = [c for c in number_cols if normalize_text(c.header) in {"ket qua", *_NUMBER_NAMES["ket qua"]}]
    return named[0] if len(named) == 1 else None


def profile_from_guess(guess: SpecGuess, fingerprint: str, source_headers: tuple[str, ...], *, confirmed_by: str, confirmed_at: str, profile_id: str, version: int = 1) -> TemplateProfile:
    """Hồ sơ từ một đoán đã có người xác nhận. Bắt buộc có người và thời điểm (Q12)."""
    if not confirmed_by or not confirmed_at:
        raise ValueError("Hồ sơ phải có người xác nhận và thời điểm")
    if guess.spec is None or guess.questions:
        raise ValueError("Còn câu hỏi chưa trả lời, chưa tạo được hồ sơ")
    spec = guess.spec
    column_map = {name: f.names[0] for name, f in spec.source_fields.items()}
    return TemplateProfile(
        profile_id=profile_id,
        version=version,
        structure_fingerprint=fingerprint,
        source_headers=source_headers,
        column_map={header: name for name, header in column_map.items()},
        dim_maps={dim: {v: v for v in values} for dim, values in spec.dim_values.items()},
        confirmed_by=confirmed_by,
        confirmed_at=confirmed_at,
    )
