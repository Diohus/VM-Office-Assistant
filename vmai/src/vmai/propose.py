"""Tạo ProposalBatch. Hàm thuần: không I/O, không mạng, không LLM, không ghi file."""
from __future__ import annotations

import datetime as dt
import uuid
from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal

from .confidence import (
    RULE_VERSION,
    CheckResult,
    Color,
    ColumnMatch,
    Confidence,
    Graded,
    ReadFacts,
    Reason,
    RowMatch,
    SourceKind,
    check_column_suspect,
    check_implausible,
    check_lot_mismatch,
    check_not_in_list,
    check_out_of_3sd,
    combine,
    grade_map,
    grade_read,
)
from .contract import (
    CellTarget,
    Change,
    ConfidenceModel,
    Evidence,
    ImageRegionLocator,
    LlmCall,
    Mapping,
    MappingMethod,
    Notice,
    PdfTextLocator,
    Previous,
    ProposalBatch,
    ReasonModel,
    Review,
    RowKeyPart,
    Source,
    Template,
    Value,
    XlsxCellLocator,
)
from .errors import FeatureNotReleased
from .headers import HeaderPlan, apply_header_map
from .mapping import match_field, normalize_text
from .phi import looks_like_person_name
from .numbers import NumberStyle, ParseStatus, infer_column_style, parse_number
from .profile import ProfileStatus
from .source_xlsx import SourceCell, SourceFacts
from .template import FormColumn, FormFacts, Limit, as_date, key_text

SUSPECT_COLUMN_SHARE = 0.5  # giả định của đội
SUSPECT_MIN_VALUES = 5
IMPLAUSIBLE_FACTOR = Decimal(5)  # Q24: tạm, cần chuyên môn và biểu mẫu thật xác nhận trước mức "Đóng"
_FORMULA_PREFIXES = ("=", "+", "-", "@")

_MATCH_ORDER = [
    ColumnMatch.NONE,
    ColumnMatch.AMBIGUOUS,
    ColumnMatch.LLM_INFERRED,
    ColumnMatch.FUZZY,
    ColumnMatch.ABBREVIATION,
    ColumnMatch.INFERRED,
    ColumnMatch.PROFILE_UNCHECKED,
    ColumnMatch.NORMALIZED,
    ColumnMatch.PROFILE,
    ColumnMatch.EXACT,
]
_METHODS = {
    ColumnMatch.EXACT: MappingMethod.EXACT,
    ColumnMatch.PROFILE: MappingMethod.PROFILE,
    ColumnMatch.PROFILE_UNCHECKED: MappingMethod.PROFILE,
    ColumnMatch.NORMALIZED: MappingMethod.NORMALIZED,
    ColumnMatch.INFERRED: MappingMethod.INFERRED,
    ColumnMatch.ABBREVIATION: MappingMethod.ABBREVIATION,
    ColumnMatch.FUZZY: MappingMethod.FUZZY,
    ColumnMatch.LLM_INFERRED: MappingMethod.LLM_INFERRED,
}


def _worst(matches: list[ColumnMatch]) -> ColumnMatch:
    return min(matches, key=_MATCH_ORDER.index)


def is_out_of_3sd(normalized: str, limit: Limit) -> bool:
    return abs(Decimal(normalized) - Decimal(str(limit.target))) > 3 * Decimal(str(limit.sd))


def plausibility_checks(normalized: str | None, limit: Limit | None) -> list[CheckResult]:
    """Q24: lệch từ 5 lần so với đích, cả hai chiều, là vô lý. Không có đích thì không chạy.

    Số ngoài miền được giữ nguyên, không tự sửa, vì QC ngoài giới hạn có thể là kết quả thật.
    """
    if normalized is None or limit is None or limit.target <= 0:
        return []
    value, target = Decimal(normalized), Decimal(str(limit.target))
    if value <= 0 or value >= target * IMPLAUSIBLE_FACTOR or value * IMPLAUSIBLE_FACTOR <= target:
        return [check_implausible(f"Số {normalized} lệch từ 5 lần so với đích {limit.target}")]
    if is_out_of_3sd(normalized, limit):
        return [check_out_of_3sd(f"Số {normalized} ngoài ±3SD quanh đích {limit.target}. Giữ nguyên, có thể là kết quả thật.")]
    return []


def _image_read(cell: SourceCell, kind: SourceKind, style: NumberStyle | None, low_quality: bool) -> tuple[ReadFacts, str | None, bool]:
    """Ô đọc từ ảnh: hai lần đọc độc lập phải cho cùng một số. So sau khi chuẩn hóa."""
    parsed = parse_number(cell.value.replace(" ", ""), style)
    read = cell.read
    second = parse_number(read.second_text, style) if read and read.second_text else None
    readers = tuple(r for r in (read.reader if read else None, read.second_reader if read else None) if r)
    agree = None
    if second is not None:
        agree = parsed.status is ParseStatus.OK and second.status is ParseStatus.OK and parsed.normalized == second.normalized
    facts = ReadFacts(
        kind,
        number_ambiguous=parsed.status is ParseStatus.AMBIGUOUS,
        not_a_number=parsed.status is ParseStatus.INVALID,
        low_quality=low_quality,
        readers=readers,
        reads_agree=agree,
    )
    return facts, parsed.normalized if parsed.status is ParseStatus.OK else None, False


def _numeric_read(cell: SourceCell, style: NumberStyle | None) -> tuple[ReadFacts, str | None, bool]:
    """Trả về sự kiện đọc, giá trị chuẩn hóa, và ô có phải số dạng chữ không."""
    is_formula = cell.formula is not None
    if is_formula and cell.kind == "empty":
        return ReadFacts(SourceKind.XLSX, is_formula=True, has_cached_value=False), None, False
    if cell.kind == "number":
        return ReadFacts(SourceKind.XLSX, is_formula=is_formula), format(Decimal(cell.value), "f"), False
    if cell.kind != "text":
        return ReadFacts(SourceKind.XLSX, not_a_number=True), None, False
    parsed = parse_number(cell.value, style)
    if parsed.status is ParseStatus.AMBIGUOUS:
        return ReadFacts(SourceKind.XLSX, number_ambiguous=True), None, True
    if parsed.status is ParseStatus.INVALID:
        return ReadFacts(SourceKind.XLSX, not_a_number=True), None, True
    return ReadFacts(SourceKind.XLSX, is_formula=is_formula), parsed.normalized, True


@dataclass(frozen=True)
class _Reading:
    """Một giá trị đọc được cho một ô đích, kèm dòng nguồn để lấy đơn vị và lô."""

    cell: SourceCell
    record: tuple[SourceCell, ...]
    match: ColumnMatch


def _force_red(confidence: Confidence, reason: Reason) -> Confidence:
    return Confidence(Color.RED, confidence.map_color, Color.RED, (*confidence.reasons, reason))


def _to_model(confidence: Confidence) -> ConfidenceModel:
    return ConfidenceModel(
        read_color=confidence.read_color,
        map_color=confidence.map_color,
        final_color=confidence.final_color,
        reasons=[ReasonModel(code=r.code, detail=r.detail) for r in confidence.reasons],
    )


def _review_unit(color: Color, has_existing: bool) -> str:
    # Ô đã có số luôn duyệt riêng, kể cả khi xanh (Q14, D6).
    if color is Color.RED or has_existing:
        return "cell"
    return "row" if color is Color.YELLOW else "group"


class _Builder:
    def __init__(self, form: FormFacts, source: SourceFacts, plan: HeaderPlan, relax_units: bool = False):
        self.relax_units = relax_units
        self.form = form
        self.spec = form.spec
        self.rows = form.row_map()
        self.source = source
        self.plan = plan
        self.idx = {name: col.index for name, col in plan.columns.items()}
        self.kind = SourceKind(source.kind)
        self.notices: list[Notice] = []
        self.changes: list[Change] = []

    def _field(self, record: tuple[SourceCell, ...], name: str | None) -> SourceCell | None:
        if name is None or name not in self.idx or self.idx[name] >= len(record):
            return None
        return record[self.idx[name]]

    def _dim_value(self, record: tuple[SourceCell, ...], dim: str) -> tuple[str, ColumnMatch] | None:
        """Giá trị chuẩn của một chiều. Mã lạ trả None để hỏi người dùng (Q16), không đoán, không gọi LLM."""
        cell = self._field(record, dim)
        if cell is None or cell.is_empty:
            return None
        if dim in self.spec.date_fields:
            day = as_date(cell.value)
            return (day.isoformat(), ColumnMatch.EXACT) if day else None
        text = cell.text.strip()
        active = self.plan.active_profile
        if active is not None and dim in active.dim_maps:
            mapped = active.dim_maps[dim].get(text)
            return (mapped, ColumnMatch.PROFILE) if mapped else None
        declared = self.spec.dim_values.get(dim)
        if declared is None:
            return key_text(cell.value), ColumnMatch.EXACT  # chiều không có bảng khai báo, ví dụ mã vật tư: khớp đúng chữ
        result = match_field(text, declared, self.spec.dim_aliases.get(dim, {}), fuzzy=False)
        if result.field is None:
            return None
        # Khớp bảng khai báo nhưng chưa có người xác nhận: máy suy ra, vàng (Q11).
        return result.field, (ColumnMatch.ABBREVIATION if result.match is ColumnMatch.ABBREVIATION else ColumnMatch.INFERRED)

    def build(self) -> ProposalBatch:
        sheet = next(s for s in self.source.sheets if s.name == self.plan.sheet)
        records = [r for r in sheet.rows[self.plan.header_row + 1 :] if any(not c.is_empty for c in r)]
        self._file_notices()

        spec = self.spec
        key_dims = [k.dim for k in spec.table.key_columns]
        col_dims = sorted({d for c in spec.table.value_columns for d in c.dims})
        grouped: dict[tuple[tuple[str, ...], str], list[_Reading]] = defaultdict(list)
        unknown: dict[str, int] = defaultdict(int)
        wrong_period = not_in_template = 0
        seen_rows: set[tuple[str, ...]] = set()

        for record in records:
            resolved = [self._dim_value(record, dim) for dim in key_dims]
            col_resolved = {dim: self._dim_value(record, dim) for dim in col_dims}
            bad_dim = next((d for d, r in col_resolved.items() if r is None), None)
            if any(r is None for r in resolved) or bad_dim is not None:
                label = self._field(record, bad_dim) if bad_dim else None
                unknown[(label.text.strip() if label else "") or "(dòng thiếu khóa)"] += 1
                continue
            key = tuple(key_text(r[0]) for r in resolved)
            if spec.period and self.form.period is not None and spec.period.field in key_dims:
                day = as_date(key[key_dims.index(spec.period.field)])
                if day and (day.year, day.month) != (self.form.period.year, self.form.period.month):
                    wrong_period += 1
                    continue
            if key not in self.rows:
                not_in_template += 1
                continue
            seen_rows.add(key)
            dims = {d: r[0] for d, r in col_resolved.items()}
            key_matches = [r[1] for r in resolved] + [self.plan.columns[n].match for n in key_dims]
            for column in self.form.columns_for(dims):
                cell = self._field(record, column.field)
                if cell is None:
                    continue
                # Chỉ tính các chiều cột này thật sự dùng: cột Người chạy không phụ thuộc mã xét nghiệm.
                matches = [*key_matches, self.plan.columns[column.field].match]
                for dim in column.dims:
                    matches += [col_resolved[dim][1], self.plan.columns[dim].match]
                grouped[(key, column.letter)].append(_Reading(cell, record, _worst(matches)))

        self._record_notices(unknown, wrong_period, not_in_template)
        missing = [k for k in self.rows if k not in seen_rows]
        if missing and not wrong_period and seen_rows:
            days = ", ".join(" ".join(k) for k in missing)
            self.notices.append(Notice(reason="SOURCE_MISSING_DAYS", count=len(missing), detail=f"Nguồn thiếu: {days}. Để trống, không tự suy số."))

        by_letter = {c.letter: c for c in self.form.columns}
        styles = self._column_styles(grouped)
        suspects = self._suspect_columns(grouped, by_letter, styles)
        for (key, letter), readings in sorted(grouped.items(), key=lambda item: (self.rows[item[0][0]][0], item[0][1])):
            self._emit(key, by_letter[letter], readings, styles.get(letter), letter in suspects)
        self.notices.extend(self._grouped_notices())

        return ProposalBatch(
            rule_version=RULE_VERSION,
            proposal_id=str(uuid.uuid4()),
            created_at=dt.datetime.now(dt.timezone.utc).isoformat(),
            template=Template(template_sha256=self.form.sha256, file_name=self.form.file_name, spec_id=spec.spec_id, spec_version=spec.version),
            sources=[Source(source_id="src-1", file_name=self.source.file_name, source_file_sha256=self.source.sha256, kind=self.kind)],
            changes=self.changes,
            notices=self.notices,
            llm_calls=list(self.plan.llm_calls),
        )

    def _file_notices(self) -> None:
        plan, source = self.plan, self.source
        add = lambda reason, count, detail: self.notices.append(Notice(reason=reason, count=count, detail=detail))  # noqa: E731
        if source.has_macros:
            add("FILE_MACRO_IGNORED", 1, "File có macro. Macro không được chạy.")
        if plan.profile_status is ProfileStatus.UNCONFIRMED:
            add("PROFILE_UNCONFIRMED", 1, "Hồ sơ chưa có người xác nhận, chưa có hiệu lực.")
        elif plan.profile_status is ProfileStatus.INVALIDATED:
            add("PROFILE_INVALIDATED", 1, "Cấu trúc biểu mẫu hoặc tiêu đề nguồn đã đổi. Hồ sơ mất hiệu lực.")
        if plan.phi_headers:
            add("PHI_COLUMN_SKIPPED", len(plan.phi_headers), f"Có cột giống dữ liệu người bệnh: {', '.join(plan.phi_headers)}. Không dùng và không gửi ra ngoài.")
        manual = [" / ".join(path) or letter for letter, path in self.form.manual_columns.items()]
        if manual or plan.person_headers:
            names = ", ".join([*manual, *plan.person_headers])
            add("PERSON_COLUMN_LEFT_BLANK", len(manual) + len(plan.person_headers), f"Thông tin về con người để nhân viên tự viết, AI không điền: {names}.")
        unchecked = [" / ".join(c.col_path) for c in self.form.columns if c.type == "number" and c.limit is None]
        if unchecked:
            add(
                "CHECK_NOT_AVAILABLE",
                len(unchecked),
                f"Biểu mẫu không có giới hạn cho: {', '.join(unchecked)}. Không có kiểm tra miền giá trị, lô, cả cột cho các cột này. Các ô này tối đa vàng.",
            )
        if plan.unused_headers:
            add("SOURCE_COLUMN_UNUSED", len(plan.unused_headers), f"Cột chưa dùng: {', '.join(plan.unused_headers)}")
        if plan.ambiguous_fields:
            add("MAP_COLUMN_AMBIGUOUS", len(plan.ambiguous_fields), f"Nhiều cột cùng là: {', '.join(plan.ambiguous_fields)}. Đang lấy cột đầu tiên, các ô liên quan là đỏ.")
        if plan.other_sheets:
            add("SOURCE_OTHER_SHEETS_IGNORED", len(plan.other_sheets), f"Sheet khác chưa đọc: {', '.join(plan.other_sheets)}")
        if source.hidden_sheets:
            add("SOURCE_HIDDEN_SHEETS_IGNORED", source.hidden_sheets, "Sheet ẩn không được đọc.")

    def _record_notices(self, unknown: dict[str, int], wrong_period: int, not_in_template: int) -> None:
        if wrong_period:
            self.notices.append(Notice(reason="SOURCE_WRONG_MONTH", count=wrong_period, detail=f"{wrong_period} dòng không thuộc kỳ của biểu mẫu. Không điền."))
        if unknown:
            detail = f"Mã chưa có trong bảng đồng nghĩa, cần hỏi người dùng: {', '.join(sorted(unknown))}"
            self.notices.append(Notice(reason="MAP_TEST_UNKNOWN", count=sum(unknown.values()), detail=detail))
        if not_in_template:
            self.notices.append(Notice(reason="MAP_ROW_NOT_FOUND", count=not_in_template, detail=f"{not_in_template} dòng có khóa không có trong biểu mẫu. Không điền."))

    def _column_styles(self, grouped) -> dict[str, NumberStyle | None]:
        texts: dict[str, list[str]] = defaultdict(list)
        for (_, letter), readings in grouped.items():
            texts[letter].extend(r.cell.value for r in readings if r.cell.is_text)
        return {letter: infer_column_style(values) for letter, values in texts.items()}

    def _suspect_columns(self, grouped, by_letter, styles) -> set[str]:
        """Nhiều số cùng ngoài ±3SD trong một cột thì nghi hồ sơ hoặc ghép cột sai."""
        counts: dict[str, list[int]] = defaultdict(lambda: [0, 0])
        for (_, letter), readings in grouped.items():
            limit = by_letter[letter].limit
            if limit is None:
                continue
            for reading in readings:
                _, normalized, _ = self._read_number(reading.cell, styles.get(letter))
                if normalized is not None:
                    counts[letter][1] += 1
                    counts[letter][0] += int(is_out_of_3sd(normalized, limit))
        return {l for l, (bad, total) in counts.items() if total >= SUSPECT_MIN_VALUES and bad / total >= SUSPECT_COLUMN_SHARE}

    def _read_number(self, cell: SourceCell, style):
        if self.kind is SourceKind.XLSX:
            return _numeric_read(cell, style)
        if self.kind is SourceKind.PDF_TEXT:
            parsed = parse_number(cell.value.replace(" ", ""), style)
            facts = ReadFacts(SourceKind.PDF_TEXT, number_ambiguous=parsed.status is ParseStatus.AMBIGUOUS, not_a_number=parsed.status is ParseStatus.INVALID)
            return facts, parsed.normalized if parsed.status is ParseStatus.OK else None, False
        return _image_read(cell, self.kind, style, self.source.low_quality)

    def _emit(self, key, column: FormColumn, readings: list[_Reading], style, suspect: bool) -> None:
        # Cùng khóa, khác giá trị: không tự chọn, hiện tất cả (X04). Cùng giá trị thì gộp một đề xuất.
        distinct: dict[str, list[_Reading]] = defaultdict(list)
        for reading in readings:
            if reading.cell.is_empty and reading.cell.formula is None:
                continue
            if column.type == "number":
                distinct[self._read_number(reading.cell, style)[1] or f"raw:{reading.cell.a1}"].append(reading)
            else:
                distinct[reading.cell.text.strip() or (reading.cell.formula or "")].append(reading)
        duplicate = len(distinct) > 1
        for group in distinct.values():
            if column.type == "number":
                self._emit_number(key, column, group, style, suspect, duplicate)
            else:
                self._emit_text(key, column, group, duplicate)

    def _target(self, key, column: FormColumn) -> CellTarget:
        table = self.spec.table
        parts = [RowKeyPart(column_path=list(self.form.key_paths[k.dim]), value=key[i]) for i, k in enumerate(table.key_columns)]
        return CellTarget(sheet=table.sheet, table_anchor=table.anchor, row_key=parts, col_path=list(column.col_path), a1_at_write=f"{column.letter}{self.rows[key][0]}")

    def _grade_mapping(self, column: FormColumn, match: ColumnMatch, key, duplicate: bool, source_unit: str | None):
        if self.spec.origin == "inferred" and match is not ColumnMatch.PROFILE:
            match = _worst([match, ColumnMatch.INFERRED])  # đoán chưa được xác nhận thì không xanh
        if match is ColumnMatch.PROFILE and not self.form.has_independent_check(column, source_unit, self.relax_units):
            match = ColumnMatch.PROFILE_UNCHECKED  # Q17
        if column.field in self.plan.ambiguous_fields:
            match = ColumnMatch.AMBIGUOUS
        if len(self.rows[key]) > 1:
            row_match = RowMatch.DUPLICATE
        elif duplicate:
            row_match = RowMatch.SOURCE_DUPLICATE
        else:
            row_match = RowMatch.KEY_EXACT
        return grade_map(match, row_match), match

    def _emit_number(self, key, column: FormColumn, group: list[_Reading], style, suspect: bool, duplicate: bool) -> None:
        reading = group[0]
        cell = reading.cell
        facts, normalized, was_text = self._read_number(cell, style)
        read = grade_read(facts)
        if was_text and normalized is not None:
            note = Reason("READ_XLSX_TEXT_NUMBER", f"Số lưu dạng chữ, đã chuẩn hóa {cell.text} thành {normalized}")
            read = Graded(read.color, (*read.reasons, note))

        limit = column.limit
        checks = plausibility_checks(normalized, limit)
        lot_cell = self._field(reading.record, self.spec.lot_field)
        source_lot = lot_cell.text.strip() if lot_cell else ""
        if limit and limit.lot and source_lot and source_lot != limit.lot:
            checks.append(check_lot_mismatch(f"Lô nguồn {source_lot}, biểu mẫu {limit.lot}"))
        if suspect:
            checks.append(check_column_suspect("Từ 50% số trong cột ngoài ±3SD. Nghi hồ sơ hoặc ghép cột sai."))

        unit_cell = self._field(reading.record, self.spec.unit_field)
        source_unit = unit_cell.text.strip() if unit_cell else ""
        mapping, match = self._grade_mapping(column, reading.match, key, duplicate, source_unit or None)
        confidence = combine(read, mapping, checks)
        if limit and limit.unit and source_unit and normalize_text(source_unit) != normalize_text(limit.unit):
            confidence = _force_red(confidence, Reason("READ_UNIT_MISMATCH", f"Đơn vị nguồn {source_unit}, biểu mẫu {limit.unit}. Không tự đổi."))
        raw = cell.formula if cell.kind == "empty" else cell.text
        value = Value(type="number", normalized=normalized or raw or "", raw_text=raw or "", unit=source_unit or None)
        self._append(key, column, value, group, confidence, match)

    def _emit_text(self, key, column: FormColumn, group: list[_Reading], duplicate: bool) -> None:
        reading = group[0]
        cell = reading.cell
        text = cell.text.strip() or (cell.formula or "")
        if looks_like_person_name(text):
            # Không điền và không chép giá trị vào thông báo.
            self.notices.append(Notice(reason="PERSON_VALUE_BLOCKED", count=1, detail=f"Ô nguồn {cell.a1} giống tên người. Không điền, để nhân viên tự viết."))
            return
        checks = []
        if column.options and text not in column.options:
            checks.append(check_not_in_list(f"{text} không có trong danh sách chọn"))
        mapping, match = self._grade_mapping(column, reading.match, key, duplicate, None)
        facts = ReadFacts(SourceKind.XLSX, is_formula=cell.formula is not None, has_cached_value=not cell.is_empty)
        confidence = combine(grade_read(facts), mapping, checks)
        if text.startswith(_FORMULA_PREFIXES):
            # Ô chữ của nguồn không được thành công thức ở biểu mẫu.
            confidence = _force_red(confidence, Reason("READ_TEXT_LOOKS_LIKE_FORMULA", "Chữ trong nguồn bắt đầu như công thức. Backend phải ghi dạng chữ."))
        self._append(key, column, Value(type="text", normalized=text, raw_text=cell.text or text), group, confidence, match)

    def _evidence(self, cell: SourceCell) -> list[Evidence]:
        read = cell.read
        if read is None:
            return [Evidence(source_id="src-1", reader="xlsx_code", read_text=cell.formula or cell.text, locator=XlsxCellLocator(sheet=self.plan.sheet, a1=cell.a1))]
        if self.kind is SourceKind.PDF_TEXT:
            return [Evidence(source_id="src-1", reader=read.reader, read_text=cell.value, locator=PdfTextLocator(page=read.page, bbox=read.bbox))]
        locator = ImageRegionLocator(page=read.page, bbox=read.bbox)
        out = [Evidence(source_id="src-1", reader=read.reader, read_text=cell.value, locator=locator)]
        if read.second_reader and read.second_text is not None:
            out.append(Evidence(source_id="src-1", reader=read.second_reader, read_text=read.second_text, locator=locator))
        return out

    def _previous(self, a1: str) -> Previous:
        existing = self.form.existing.get(a1)
        if existing is None:
            return Previous()
        if existing.formula is not None:
            return Previous(formula=existing.formula)
        kind = "number" if existing.is_number else "text"
        return Previous(value=Value(type=kind, normalized=existing.value or "", raw_text=existing.value or ""))

    def _append(self, key, column: FormColumn, value: Value, group: list[_Reading], confidence: Confidence, match: ColumnMatch) -> None:
        target = self._target(key, column)
        previous = self._previous(target.a1_at_write)
        active = self.plan.active_profile
        evidence = [e for r in group for e in self._evidence(r.cell)]
        self.changes.append(
            Change(
                change_id=f"c-{len(self.changes) + 1}",
                op="set_value",
                target=target,
                previous=previous,
                value=value,
                evidence=evidence,
                mapping=Mapping(
                    method=_METHODS.get(match, MappingMethod.NONE),
                    source_header=self.plan.columns[column.field].header,
                    profile_id=active.profile_id if active else None,
                    profile_version=active.version if active else None,
                ),
                confidence=_to_model(confidence),
                review=Review(unit=_review_unit(confidence.final_color, previous.value is not None or previous.formula is not None)),
            )
        )

    def _grouped_notices(self) -> list[Notice]:
        grouped: dict[str, list[str]] = defaultdict(list)
        details: dict[str, str] = {}
        prefixes = ("CHECK_", "READ_UNIT", "READ_NUMBER", "READ_NOT", "READ_TEXT", "READ_FORMULA", "SOURCE_DUPLICATE")
        for change in self.changes:
            for reason in change.confidence.reasons:
                if reason.code.startswith(prefixes):
                    grouped[reason.code].append(change.change_id)
                    details.setdefault(reason.code, reason.detail)
        return [Notice(reason=c, count=len(ids), change_ids=ids, detail=details[c]) for c, ids in grouped.items()]


def propose(
    form: FormFacts,
    source: SourceFacts,
    plan: HeaderPlan,
    header_map: dict[str, str | None] | None = None,
    llm_calls: tuple[LlmCall, ...] = (),
    *,
    measure_relax_units: bool = False,  # chỉ để đo Q30, xem has_independent_check
    experimental_image_sources: bool = False,  # Q81: đọc ảnh, PDF thành bảng chưa phát hành, chỉ bật khi đo
) -> ProposalBatch:
    """Thiếu cột bắt buộc thì ném `MissingFields`, chưa có đề xuất nào được tạo.

    Cùng đầu vào cho cùng đề xuất, trừ `proposal_id` và `created_at`.
    """
    if source.kind != "xlsx" and not experimental_image_sources:
        raise FeatureNotReleased("Điền từ ảnh và PDF")
    return _Builder(form, source, apply_header_map(plan, source, header_map, llm_calls), measure_relax_units).build()
