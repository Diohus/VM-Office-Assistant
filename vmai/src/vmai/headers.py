"""Ghép tiêu đề cột của nguồn. Thuần, không mạng: `plan_headers` nói cần hỏi LLM gì, backend tự gọi dịch vụ `llm`.

Chỉ tiêu đề đi qua ranh giới này (Q16). `SourceFacts` và `FormFacts` không đi vào `llm`.
"""
from __future__ import annotations

from pydantic import BaseModel, ConfigDict

from .confidence import ColumnMatch
from .contract import LlmCall
from .errors import MissingFields, SourceFileError
from .llm_tasks import sendable_headers
from .mapping import match_field
from .phi import is_person_header, phi_kind
from .profile import ProfileStatus, TemplateProfile, profile_status
from .source_xlsx import SourceFacts, SourceSheet
from .template import FormFacts, TemplateSpec

HEADER_SEARCH_ROWS = 20


class ResolvedColumn(BaseModel):
    model_config = ConfigDict(frozen=True)

    index: int
    match: ColumnMatch
    header: str


class HeaderPlan(BaseModel):
    """Kết quả ghép bằng code. `llm_headers` rỗng nghĩa là không cần gọi LLM."""

    model_config = ConfigDict(frozen=True)

    sheet: str
    header_row: int
    headers: tuple[str, ...]
    columns: dict[str, ResolvedColumn]
    # Tên cột chưa dùng và cột giống dữ liệu người bệnh. Giá trị của chúng không đi tiếp.
    unused_headers: tuple[str, ...] = ()
    phi_headers: tuple[str, ...] = ()
    person_headers: tuple[str, ...] = ()
    ambiguous_fields: tuple[str, ...] = ()
    profile_status: ProfileStatus = ProfileStatus.NONE
    active_profile: TemplateProfile | None = None
    other_sheets: tuple[str, ...] = ()
    missing_fields: tuple[str, ...] = ()
    llm_headers: tuple[str, ...] = ()
    llm_fields: dict[str, str] = {}
    llm_calls: tuple[LlmCall, ...] = ()


def _classify(row, spec: TemplateSpec, profile: TemplateProfile | None):
    fields = {name: f.names for name, f in spec.source_fields.items()}
    aliases = {name: f.aliases for name, f in spec.source_fields.items()}
    columns: dict[str, ResolvedColumn] = {}
    unused: list[str] = []
    phi: list[str] = []
    people: list[str] = []
    ambiguous: list[str] = []
    for index, cell in enumerate(row):
        if not cell.is_text or cell.is_empty:
            continue
        header = cell.value
        # Tiêu đề giống dữ liệu người bệnh bị loại trước mọi bước khác, kể cả hồ sơ.
        if phi_kind(header):
            phi.append(header)
            continue
        if is_person_header(header):
            people.append(header)
            continue
        if profile is not None and header in profile.column_map:
            name, match = profile.column_map[header], ColumnMatch.PROFILE
        else:
            result = match_field(header, fields, aliases)
            name, match = result.field, result.match
        if name is None:
            unused.append(header)
        elif name in columns:
            ambiguous.append(name)
        else:
            columns[name] = ResolvedColumn(index=index, match=match, header=header)
    return columns, unused, phi, people, ambiguous


def _pick_table(source: SourceFacts, spec: TemplateSpec, profile: TemplateProfile | None):
    best = None
    for sheet in source.sheets:
        for index, row in enumerate(sheet.rows[:HEADER_SEARCH_ROWS]):
            columns, unused, phi, people, ambiguous = _classify(row, spec, profile)
            if columns and (best is None or len(columns) > len(best[2])):
                best = (sheet, index, columns, unused, phi, people, ambiguous)
    return best


def _headers_in(sheet: SourceSheet, index: int) -> list[str]:
    return [c.value for c in sheet.rows[index] if c.is_text and not c.is_empty]


def plan_headers(form: FormFacts, source: SourceFacts, profile: TemplateProfile | None = None) -> HeaderPlan:
    """Ghép bằng hồ sơ và bảng khai báo. Chỉ đề nghị hỏi LLM khi còn thiếu cột bắt buộc."""
    spec = form.spec
    best = _pick_table(source, spec, None)
    if best is None:
        raise SourceFileError("NO_MATCHING_TABLE", "Không tìm thấy bảng kết quả phù hợp trong file")

    # Đọc tiêu đề trước để biết hồ sơ còn hiệu lực không, rồi mới dùng hồ sơ.
    usable = [h for h in _headers_in(best[0], best[1]) if not phi_kind(h) and not is_person_header(h)]
    status = profile_status(profile, form.fingerprint, usable)
    active = profile if status is ProfileStatus.ACTIVE else None
    if active is not None:
        best = _pick_table(source, spec, active) or best
    sheet, index, columns, unused, phi, people, ambiguous = best

    required = [name for name, f in spec.source_fields.items() if f.required]
    missing = tuple(name for name in required if name not in columns)
    llm_headers = tuple(sendable_headers(unused)) if missing else ()
    llm_fields = {name: f.names[0] for name, f in spec.source_fields.items() if name not in columns} if llm_headers else {}
    return HeaderPlan(
        sheet=sheet.name,
        header_row=index,
        headers=tuple(_headers_in(sheet, index)),
        columns=columns,
        unused_headers=tuple(unused),
        phi_headers=tuple(phi),
        person_headers=tuple(people),
        ambiguous_fields=tuple(ambiguous),
        profile_status=status,
        active_profile=active,
        other_sheets=tuple(s.name for s in source.sheets if s.name != sheet.name),
        missing_fields=missing,
        llm_headers=llm_headers,
        llm_fields=llm_fields,
    )


def apply_header_map(plan: HeaderPlan, source: SourceFacts, header_map: dict[str, str | None] | None, llm_calls: tuple[LlmCall, ...] = ()) -> HeaderPlan:
    """Áp kết quả LLM vào kế hoạch. Chỉ nhận tiêu đề đã hỏi và trường đã đề nghị, nên màu tối đa là vàng."""
    columns = dict(plan.columns)
    unused = list(plan.unused_headers)
    if header_map:
        row = next(s for s in source.sheets if s.name == plan.sheet).rows[plan.header_row]
        positions = {c.value: i for i, c in enumerate(row) if c.is_text}
        for header, name in header_map.items():
            if header not in plan.llm_headers or name is None or name not in plan.llm_fields or name in columns:
                continue
            columns[name] = ResolvedColumn(index=positions[header], match=ColumnMatch.LLM_INFERRED, header=header)
            unused.remove(header)
    missing = tuple(name for name in plan.missing_fields if name not in columns)
    if missing:
        raise MissingFields(list(missing), list(plan.headers))
    return plan.model_copy(update={"columns": columns, "unused_headers": tuple(unused), "missing_fields": (), "llm_calls": llm_calls})
