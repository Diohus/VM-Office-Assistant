"""Dữ liệu cho màn "Xác nhận cách ghép" (US-1.13): cột đã ghép, cột còn trống, giá trị ví dụ, các cột nguồn để chọn.

Hàm thuần. Giá trị ví dụ chỉ để hiện cho người dùng, không ghi log và không gửi LLM. Giá trị giống tên người bị bỏ.
"""
from __future__ import annotations

from pydantic import BaseModel, ConfigDict

from .infer import SpecGuess, _source_columns
from .phi import looks_like_person_name
from .source_xlsx import SourceFacts
from .template import key_text

SAMPLE_COUNT = 3


class ConfirmRow(BaseModel):
    model_config = ConfigDict(frozen=True)

    letter: str
    col_path: tuple[str, ...]
    source_header: str | None  # None: chưa ghép được, người dùng phải chọn hoặc để trống
    samples: tuple[str, ...] = ()
    dims: dict[str, str] = {}
    chosen_by_user: bool = False


class SourceOption(BaseModel):
    model_config = ConfigDict(frozen=True)

    header: str
    samples: tuple[str, ...] = ()
    usable: bool = True  # False: hai cột cùng tên, chưa phân biệt được khi chọn


def _samples(table, column, dims: dict[str, str], count: int) -> tuple[str, ...]:
    """Giá trị đầu tiên của cột, đúng các dòng thuộc chiều đó. Dạng dọc mà lấy cả cột thì ví dụ sẽ lẫn xét nghiệm khác."""
    out: list[str] = []
    for r in table.data_rows:
        match = True
        for field, wanted in dims.items():
            index = int(field[1:]) + 1  # tên trường do infer_spec đặt là f<chỉ số cột>
            cell = table.grid.own(r, index)
            if cell is None or key_text(cell.value) != key_text(wanted):
                match = False
                break
        if not match:
            continue
        cell = table.grid.own(r, column.index + 1)
        if cell is None or cell.is_empty or looks_like_person_name(cell.text):
            continue
        out.append(cell.text.strip())
        if len(out) == count:
            break
    return tuple(out)


def confirmation_rows(guess: SpecGuess, source: SourceFacts, sample_count: int = SAMPLE_COUNT) -> list[ConfirmRow]:
    """Một dòng cho mỗi cột đích: cột đã ghép kèm giá trị ví dụ, rồi cột chưa ghép được."""
    columns, table, _ = _source_columns(source)
    rows = []
    for item in guess.columns:
        found = next((c for c in columns if c.header == item.source_header), None)
        samples = _samples(table, found, item.dims, sample_count) if found is not None and table is not None else ()
        rows.append(ConfirmRow(letter=item.letter, col_path=item.col_path, source_header=item.source_header, samples=samples, dims=item.dims, chosen_by_user=item.chosen_by_user))
    rows.extend(ConfirmRow(letter=item.letter, col_path=item.col_path, source_header=None) for item in guess.open_columns)
    return sorted(rows, key=lambda r: (len(r.letter), r.letter))


def source_options(source: SourceFacts, sample_count: int = SAMPLE_COUNT) -> list[SourceOption]:
    """Các cột số của nguồn người dùng có thể chọn. Cột về người, dữ liệu người bệnh, cột ngày không có ở đây."""
    columns, table, _ = _source_columns(source)
    numbers = [c for c in columns if c.is_number and not c.is_date]
    names = [c.header for c in numbers]
    return [SourceOption(header=c.header, samples=_samples(table, c, {}, sample_count), usable=names.count(c.header) == 1) for c in numbers]
