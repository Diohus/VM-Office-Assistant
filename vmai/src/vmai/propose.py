"""Tạo ProposalBatch. Hàm thuần: không I/O, không mạng, không LLM, không ghi file."""
from __future__ import annotations

import datetime as dt
import re
import statistics
import uuid
from collections import Counter, defaultdict
from collections.abc import Mapping as _Mapping
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation

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
    check_decimal_places,
    check_decimal_shift,
    check_implausible,
    check_lot_mismatch,
    check_not_in_list,
    check_out_of_3sd,
    check_printed_total,
    check_total_not_comparable,
    check_unit_one_side,
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
from .mapping import match_field, normalize_text, same_unit, split_header_unit
from .phi import looks_like_person_name, phi_kind
from .numbers import NumberStyle, ParseStatus, infer_column_style, parse_number
from .profile import ProfileStatus
from .source_xlsx import SourceCell, SourceFacts
from .template import FormColumn, FormFacts, Limit, as_date, classify_text, key_text

SUSPECT_COLUMN_SHARE = 0.5  # giả định của đội
SUSPECT_MIN_VALUES = 5
IMPLAUSIBLE_FACTOR = Decimal(5)  # Q24: tạm, cần chuyên môn và biểu mẫu thật xác nhận trước mức "Đóng"
_FORMULA_PREFIXES = ("=", "+", "-", "@")
DECIMAL_MIN_VALUES = 5  # giả định của đội, như SUSPECT_MIN_VALUES
DECIMAL_MIN_SHARE = 0.6  # giả định: cách viết phổ biến của cột phải chiếm từ 60%
DECIMAL_MIN_REFERENCE = 3  # giả định: tối thiểu số ô làm mốc
NARROW_COLUMN_RATIO = Decimal(3)  # giả định: cột có tứ phân vị trên/dưới từ 3 trở xuống mới coi là cùng cỡ số
_TOTAL_LINE = re.compile(r"^tong(?: cong| so(?: mau| ket qua| dong)?)? (\d+)$")
_UNREAD_ADVICE = "Có thể app đọc thiếu, phiếu in chưa chắc sai. Nên chụp lại ảnh rõ hơn, hoặc đối chiếu từng ô với ảnh trước khi duyệt."
_MISMATCH_ADVICE = "Không khớp với số ghi trên phiếu. Nên xem lại phiếu gốc: có thể phiếu ghi sai, hoặc có dòng app không thấy."
_TOTAL_KEYS = frozenset({"tong", "tong cong", "cong", "total", "grand total"})

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


# Q119, T-AI-14: nghi nạp nhầm file tháng trước. Ngưỡng đặt trên phân bố tự nghĩ, chưa có số thật.
STALE_ROW_SHARE = 0.30
STALE_CONSTANT_SHARE = 0.80
STALE_MIN_CHANGING_COLUMNS = 2
STALE_MIN_SAME_DAYS = 2  # Q120: lô 3 ngày, 1 ngày trùng tự nhiên đã là 33%; cần ít nhất 2 ngày trùng


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
        return [check_implausible(f"Số {normalized} gấp từ 5 lần đích {limit.target} trở lên, hoặc chỉ bằng từ 1/5 đích trở xuống, hoặc không lớn hơn 0")]
    if is_out_of_3sd(normalized, limit):
        return [check_out_of_3sd(f"Số {normalized} ngoài ±3SD quanh đích {limit.target}. Giữ nguyên, có thể là kết quả thật.")]
    return []


def _image_read(cell: SourceCell, kind: SourceKind, style: NumberStyle | None, source: SourceFacts) -> tuple[ReadFacts, str | None, bool]:
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
        low_quality=source.low_quality,
        handwriting_declared=source.handwriting_declared,
        scanned_text_layer=source.scanned_text_layer,
        readers=readers,
        reads_agree=agree,
    )
    return facts, parsed.normalized if parsed.status is ParseStatus.OK else None, False


def _same_key(cell: SourceCell, second_text: str | None) -> bool:
    """Khóa đọc từ ảnh: lần đọc thứ hai phải ra cùng một khóa. So ngày theo ngày, chữ sau khi bỏ dấu và khoảng trắng."""
    if not second_text:
        return False
    kind, value = classify_text(second_text)
    if cell.kind == "date" or kind == "date":
        return cell.kind == "date" and kind == "date" and cell.value == value
    return normalize_text(cell.value).replace(" ", "") == normalize_text(value).replace(" ", "")


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
    key_confirmed: bool = True


def _day_label(key: tuple[str, ...]) -> str:
    """Khóa dòng thành chữ ngắn để nêu trong câu: ngày ISO thành dd/mm, khóa khác giữ nguyên."""
    day = as_date(key[0]) if key else None
    return day.strftime("%d/%m") if day else " ".join(key)


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



UNKNOWN_NAMES_SHOWN = 3
_MAX_NAME_LENGTH = 30


def _safe_test_name(name: str) -> bool:
    """Tên lạ ở cột khóa chỉ được hiện khi trông như tên xét nghiệm. Giống tên người, mã người bệnh, chuỗi dài hay nhiều số
    thì không hiện (cột khóa có thể bị lẫn dữ liệu khác), chỉ đếm."""
    digits = sum(c.isdigit() for c in name)
    return 0 < len(name) <= _MAX_NAME_LENGTH and digits <= 3 and not looks_like_person_name(name) and phi_kind(name) is None


def _unknown_names_detail(unknown: dict[str, int]) -> str:
    """Tối đa 3 tên an toàn, rồi "và N tên khác". Không bao giờ có tên giống thông tin về người."""
    names = sorted(unknown)
    safe = [n for n in names if _safe_test_name(n)]
    shown = safe[:UNKNOWN_NAMES_SHOWN]
    rest = len(names) - len(shown)
    hidden = len(safe) < len(names)  # có tên bị giấu: nói rõ là không hiện ở đây (PO/BA 11/11)
    if not shown:
        return f"{rest} tên, app không hiện ở đây"
    if not rest:
        return ", ".join(shown)
    return ", ".join(shown) + f" và {rest} tên khác" + (" không hiện ở đây" if hidden else "")


class _Builder:
    def __init__(self, form: FormFacts, source: SourceFacts, plan: HeaderPlan, relax_units: bool = False, previous_month: _Mapping[str, str] | None = None):
        self.relax_units = relax_units
        self.previous_month = previous_month
        self.form = form
        self.spec = form.spec
        self.rows = form.row_map()
        self.source = source
        self.plan = plan
        self.idx = {name: col.index for name, col in plan.columns.items()}
        self.kind = SourceKind(source.kind)
        self.notices: list[Notice] = []
        self.changes: list[Change] = []
        self.total_all: CheckResult | None = None
        self.total_columns: dict[str, CheckResult] = {}
        self.decimal_checks: dict[str, CheckResult] = {}

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

    def _key_confirmed(self, record: tuple[SourceCell, ...], dims: list[str]) -> bool:
        """Khóa đọc từ ảnh cũng phải qua hai lần đọc. Khóa lệch thì số có thể vào nhầm dòng (MAP_ROW_KEY_UNCONFIRMED)."""
        if self.kind not in (SourceKind.IMAGE, SourceKind.PDF_IMAGE):
            return True
        for dim in dims:
            cell = self._field(record, dim)
            if cell is None or cell.read is None or not _same_key(cell, cell.read.second_text):
                return False
        return True

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
                confirmed = self._key_confirmed(record, [*key_dims, *column.dims])
                grouped[(key, column.letter)].append(_Reading(cell, record, _worst(matches), confirmed))

        self._record_notices(unknown, wrong_period, not_in_template)
        missing = [k for k in self.rows if k not in seen_rows]
        if missing and not wrong_period and seen_rows:
            days = ", ".join(" ".join(k) for k in missing)
            self.notices.append(Notice(reason="SOURCE_MISSING_DAYS", count=len(missing), detail=f"Nguồn thiếu: {days}. Để trống, không tự suy số."))

        by_letter = {c.letter: c for c in self.form.columns}
        styles = self._column_styles(grouped)
        suspects = self._suspect_columns(grouped, by_letter, styles)
        self.decimal_checks = self._decimal_findings(grouped, by_letter, styles)
        self._printed_totals(sheet, records, key_dims, grouped, by_letter, styles)
        for (key, letter), readings in sorted(grouped.items(), key=lambda item: (self.rows[item[0][0]][0], item[0][1])):
            self._emit(key, by_letter[letter], readings, styles.get(letter), letter in suspects)
        if self.previous_month:
            self._same_as_last_month(self.previous_month)
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

    def _same_as_last_month(self, previous: _Mapping[str, str]) -> None:
        """Q119, T-AI-14: nghi nạp nhầm file tháng trước. Chỉ hạ màu, không bao giờ nâng.

        So theo dòng: một dòng giống hệt khi mọi cột đang thay đổi có cùng số với tháng trước ở cùng ô.
        Bỏ cột gần như cố định (một giá trị chiếm từ 80% số ô của cột đó ở tháng trước).
        Q121: dòng giống hệt thì chính dòng đó tối đa vàng, không ngưỡng, kể cả khi chỉ 1-2 cột thay đổi.
        Q119, Q120: có từ 3 cột thay đổi, quá 30% số dòng so được giống hệt và ít nhất 2 dòng thì cả lô tối đa vàng, kèm thông báo. Ngưỡng đặt trên phân bố tự nghĩ (eval/run_stale_share.py).
        """
        def number(text: str | None) -> Decimal | None:
            try:
                return Decimal(str(text)) if text not in (None, "") else None
            except InvalidOperation:
                return None

        cells: dict[tuple[str, int], Decimal] = {}
        for change in self.changes:
            if change.value is None or change.value.type != "number":
                continue
            a1 = change.target.a1_at_write
            letter = a1.rstrip("0123456789")
            if (value := number(change.value.normalized)) is not None:
                cells[(letter, int(a1[len(letter):]))] = value
        last: dict[tuple[str, int], Decimal] = {}
        for a1, text in previous.items():
            letter = a1.rstrip("0123456789")
            if letter and a1[len(letter):].isdigit() and (value := number(text)) is not None:
                last[(letter, int(a1[len(letter):]))] = value
        letters = sorted({letter for letter, _ in cells})
        changing = []
        for letter in letters:
            column = [v for (l, _), v in last.items() if l == letter]
            if column and Counter(column).most_common(1)[0][1] < STALE_CONSTANT_SHARE * len(column):
                changing.append(letter)
        if not changing:
            return
        rows = sorted({row for _, row in cells})
        compared = [row for row in rows if all((l, row) in cells and (l, row) in last for l in changing)]
        same = [row for row in compared if all(cells[(l, row)] == last[(l, row)] for l in changing)]
        if not same:
            return
        label_of = {row: _day_label(key) for key, form_rows in self.rows.items() for row in form_rows}

        # Q121: chính dòng trùng hệt tháng trước tối đa vàng, không ngưỡng, mọi số cột thay đổi. Không thông báo chung.
        row_reason = {
            row: ReasonModel(code="CHECK_ROW_SAME_AS_LAST_MONTH",
                             detail=f"Ngày {label_of.get(row, f'dòng {row}')} giống hệt tháng trước ở mọi cột kết quả. Có thể đây là số của tháng trước.")
            for row in same
        }
        # Q119, Q120: cả lô tối đa vàng, kèm thông báo chung, khi đủ điều kiện ngưỡng.
        batch_reason = None
        if len(changing) > STALE_MIN_CHANGING_COLUMNS and len(same) / len(compared) > STALE_ROW_SHARE and len(same) >= STALE_MIN_SAME_DAYS:
            # T-AI-15: nêu ngày cụ thể (chỉ ngày, không bao giờ giá trị ô), nói về file chứ không đổ lỗi cho người dùng.
            days = [label_of.get(row, f"dòng {row}") for row in same]
            listed = f"Ngày {', '.join(days)}" if len(days) <= 3 else f"{len(days)} ngày (ví dụ {', '.join(days[:3])})"
            detail = (f"{listed} giống hệt tháng trước ở mọi cột kết quả (trong {len(compared)} ngày có số tháng trước để so). "
                      "Có thể đây là file của tháng trước. Mở file nguồn, xem ngày tháng ở đầu file. Nếu đúng file tháng này, duyệt như thường.")
            batch_reason = ReasonModel(code="CHECK_SAME_AS_LAST_MONTH", detail=detail)
            self.notices.append(Notice(reason="SOURCE_LOOKS_LIKE_LAST_MONTH", count=len(same), detail=detail))
        updated = []
        for change in self.changes:
            if change.value is None or change.value.type != "number":
                updated.append(change)
                continue
            a1 = change.target.a1_at_write
            row = int(a1[len(a1.rstrip("0123456789")):])
            added = [r for r in (row_reason.get(row), batch_reason) if r is not None]
            if not added:
                updated.append(change)
                continue
            conf = change.confidence
            final = Color.YELLOW if conf.final_color is Color.GREEN else conf.final_color
            has_previous = change.previous.value is not None or change.previous.formula is not None
            updated.append(change.model_copy(update={
                "confidence": conf.model_copy(update={"final_color": final, "reasons": [*conf.reasons, *added]}),
                "review": change.review.model_copy(update={"unit": _review_unit(final, has_previous)}),
            }))
        self.changes = updated

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
            detail = _unknown_names_detail(unknown)  # giao diện điền vào {tên} của câu hiện
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

    def _decimal_findings(self, grouped, by_letter, styles) -> dict[str, CheckResult]:
        """Số chữ số thập phân khác cả cột, và dấu thập phân bị mất hoặc dịch trên ảnh.

        Ô số thật của Excel không xét, vì kiểu số không giữ chữ số 0 cuối. Hai lần đọc lệch thì đã đỏ và không dùng làm mốc.
        Hai engine cùng mất dấu vẫn "đồng ý" nên R9 không bắt được: phải so với các ô khác trong cột.
        Cả cột cùng mất dấu thì không có mốc để so, kiểm tra này không bắt được.
        """
        out: dict[str, CheckResult] = {}
        is_image = self.kind in (SourceKind.IMAGE, SourceKind.PDF_IMAGE)
        columns: dict[str, list[tuple[SourceCell, Decimal, int]]] = defaultdict(list)
        for (_, letter), readings in grouped.items():
            if by_letter[letter].type != "number":
                continue
            for reading in readings:
                if not reading.cell.is_text:
                    continue
                facts, normalized, _ = self._read_number(reading.cell, styles.get(letter))
                if normalized is None or (is_image and not facts.reads_agree):
                    continue
                columns[letter].append((reading.cell, Decimal(normalized), len(normalized.partition(".")[2])))
        for letter, items in columns.items():
            if len(items) < DECIMAL_MIN_VALUES:
                continue
            counts: dict[int, int] = defaultdict(int)
            for _, _, places in items:
                counts[places] += 1
            usual = max(counts, key=counts.get)
            if counts[usual] / len(items) < DECIMAL_MIN_SHARE:
                continue
            for cell, _, places in items:
                if places != usual:
                    out[cell.a1] = check_decimal_places(f"Ô có {places} chữ số thập phân, các ô cùng cột thường có {usual}.")
            if self.kind is not SourceKind.XLSX:  # PDF có thể xanh nên càng phải xét
                out.update(self._decimal_shifts(items, usual))
        return out

    @staticmethod
    def _decimal_shifts(items: list[tuple[SourceCell, Decimal, int]], usual: int) -> dict[str, CheckResult]:
        positive = [i for i in items if i[1] > 0]
        reference = sorted(v for _, v, places in positive if places == usual)
        if len(reference) < DECIMAL_MIN_REFERENCE:
            return {}
        middle = Decimal(statistics.median(reference))
        low, high = reference[len(reference) // 4], reference[-(len(reference) // 4) - 1]
        # Cột toàn số nguyên không có dấu để dịch. Số lệch cỡ ở đó thường là số thật, như một món hàng đắt hơn hẳn.
        narrow = usual > 0 and low > 0 and high / low <= NARROW_COLUMN_RATIO
        odd, more_places = [], False
        for item in positive:
            cell, value, places = item
            if max(value / middle, middle / value) >= IMPLAUSIBLE_FACTOR and (places != usual or narrow):
                odd.append(cell)
                more_places = more_places or places > usual
        # Ô lệch lại có nhiều chữ số thập phân hơn: có thể chính số đông đã mất dấu, không biết bên nào đúng nên hạ cả cột.
        flagged = [i[0] for i in positive] if more_places else odd
        detail = "Cỡ số lệch từ 5 lần so với các ô cùng cột dù hai lần đọc giống nhau. Có thể cả hai cùng mất hoặc dịch dấu thập phân."
        return {cell.a1: check_decimal_shift(detail) for cell in flagged}

    def _printed_totals(self, sheet, records, key_dims, grouped, by_letter, styles) -> None:
        """Dòng tổng in trên phiếu. Chỉ hạ màu; khớp không nâng màu ô nào. Đọc không ra thì không kết luận."""
        sums: dict[str, Decimal | None] = {}
        filled = 0
        filled_by_letter: dict[str, int] = defaultdict(int)
        for (_, letter), readings in grouped.items():
            if by_letter[letter].type != "number":
                continue
            for reading in readings:
                if reading.cell.is_empty:
                    continue
                filled += 1
                filled_by_letter[letter] += 1
                _, normalized, _ = self._read_number(reading.cell, styles.get(letter))
                current = sums.get(letter, Decimal(0))
                sums[letter] = None if normalized is None or current is None else current + Decimal(normalized)
        data_rows = [r for r in records if not self._is_total_row(r, key_dims)]
        # Dòng tổng theo cột: `Tổng | 12,5 | 30`.
        for record in records:
            if not self._is_total_row(record, key_dims):
                continue
            for letter, total in sums.items():
                cell = self._field(record, by_letter[letter].field)
                if total is None or cell is None or cell.is_empty:
                    continue
                _, normalized, _ = self._read_number(cell, styles.get(letter))
                if normalized is None:
                    continue
                printed = Decimal(normalized)
                if abs(total - printed) > Decimal(5) / 10 ** (len(normalized.partition(".")[2]) + 1):
                    empty = len(self.rows) - filled_by_letter[letter]
                    if empty > 0:
                        self.total_columns[letter] = check_total_not_comparable(f"Phiếu ghi tổng {normalized}, nhưng cột này còn {empty} ô app chưa đọc được nên chưa so được. {_UNREAD_ADVICE}")
                    else:
                        self.total_columns[letter] = check_printed_total(f"Phiếu ghi tổng {normalized}, cộng các ô app đọc được ra {total}. {_MISMATCH_ADVICE}")
        # Dòng đếm: `Tổng số mẫu: 187`. Phiếu có thể đếm theo dòng hoặc theo kết quả.
        lines = list(sheet.printed_lines) or [" ".join(c.text for c in row if not c.is_empty) for row in sheet.rows]
        accepted = {len(data_rows), filled, *(int(s) for s in sums.values() if s is not None and s == int(s))}
        for line in lines:
            match = _TOTAL_LINE.match(normalize_text(line))
            if match and int(match.group(1)) not in accepted:
                printed = int(match.group(1))
                number_columns = sum(c.type == "number" for c in by_letter.values())
                empty_cells = len(self.rows) * number_columns - filled
                missing_rows = len(self.rows) - len(data_rows)
                if (filled < printed <= filled + empty_cells) or (len(data_rows) < printed <= len(data_rows) + missing_rows):
                    self.total_all = check_total_not_comparable(f"Phiếu ghi tổng {printed}, app mới đọc được {filled} ô, còn {empty_cells} ô chưa đọc được nên chưa so được. {_UNREAD_ADVICE}")
                else:
                    self.total_all = check_printed_total(f"Phiếu ghi tổng {printed}, app đếm được {len(data_rows)} dòng, {filled} ô. {_MISMATCH_ADVICE}")
                break

    def _is_total_row(self, record: tuple[SourceCell, ...], key_dims: list[str]) -> bool:
        for dim in key_dims:
            cell = self._field(record, dim)
            if cell is not None and not cell.is_empty and normalize_text(cell.text) in _TOTAL_KEYS:
                return True
        return False

    def _read_number(self, cell: SourceCell, style):
        if self.kind is SourceKind.XLSX:
            return _numeric_read(cell, style)
        if self.kind is SourceKind.PDF_TEXT:
            parsed = parse_number(cell.value.replace(" ", ""), style)
            facts = ReadFacts(SourceKind.PDF_TEXT, number_ambiguous=parsed.status is ParseStatus.AMBIGUOUS, not_a_number=parsed.status is ParseStatus.INVALID, pdf_text_verified=self.source.pdf_text_verified)
            return facts, parsed.normalized if parsed.status is ParseStatus.OK else None, False
        return _image_read(cell, self.kind, style, self.source)

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
        start = len(self.changes)
        for group in distinct.values():
            if column.type == "number":
                self._emit_number(key, column, group, style, suspect, duplicate)
            else:
                self._emit_text(key, column, group, duplicate)
        if duplicate and len(self.changes) - start > 1:
            self._merge_duplicates(start)

    def _merge_duplicates(self, start: int) -> None:
        """Một ô chỉ có một đề xuất (D10). Nhiều lần đọc khác số thì gộp thành một Change đỏ, không chọn số nào
        (`value` rỗng), `evidence` giữ mọi lần đọc kèm số đã chuẩn hóa để người dùng chọn (X04)."""
        parts = self.changes[start:]
        del self.changes[start:]
        first = parts[0]
        evidence = [e.model_copy(update={"normalized": c.value.normalized if c.value else None}) for c in parts for e in c.evidence]
        reasons, seen = [], set()
        for c in parts:
            for r in c.confidence.reasons:
                if (r.code, r.detail) not in seen:
                    seen.add((r.code, r.detail))
                    reasons.append(r)
        ranks = {Color.RED: 0, Color.YELLOW: 1, Color.GREEN: 2}
        lowest = lambda attr: min((getattr(c.confidence, attr) for c in parts), key=ranks.__getitem__)  # noqa: E731
        confidence = ConfidenceModel(read_color=lowest("read_color"), map_color=lowest("map_color"), final_color=Color.RED, reasons=reasons)
        merged = first.model_copy(update={"value": None, "evidence": evidence, "confidence": confidence, "review": first.review.model_copy(update={"unit": "cell"})})
        self.changes.append(Change.model_validate(merged.model_dump()))

    def _target(self, key, column: FormColumn) -> CellTarget:
        table = self.spec.table
        parts = [RowKeyPart(column_path=list(self.form.key_paths[k.dim]), value=key[i]) for i, k in enumerate(table.key_columns)]
        return CellTarget(sheet=table.sheet, table_anchor=table.anchor, row_key=parts, col_path=list(column.col_path), a1_at_write=f"{column.letter}{self.rows[key][0]}")

    def _grade_mapping(self, column: FormColumn, match: ColumnMatch, key, duplicate: bool, source_unit: str | None, key_confirmed: bool = True):
        if self.spec.origin == "inferred" and match is not ColumnMatch.PROFILE:
            match = _worst([match, ColumnMatch.INFERRED])  # đoán chưa được xác nhận thì không xanh
        if match is ColumnMatch.PROFILE and (column.field in self.spec.user_chosen_fields or column.letter in self.spec.level_answer_columns):
            match = ColumnMatch.PROFILE_USER_ANSWER  # Q114, Q115: tối đa vàng mãi, không chờ Q17
        elif match is ColumnMatch.PROFILE and column.field in self.spec.alternative_value_fields:
            match = ColumnMatch.PROFILE_SIMILAR_COLUMNS  # T-AI-12: bộ đoán đã chọn giữa nhiều cột số
        if match is ColumnMatch.PROFILE and not self.form.has_independent_check(column, source_unit, self.relax_units):
            match = ColumnMatch.PROFILE_UNCHECKED  # Q17
        if column.field in self.plan.ambiguous_fields:
            match = ColumnMatch.AMBIGUOUS
        if len(self.rows[key]) > 1:
            row_match = RowMatch.DUPLICATE
        elif duplicate:
            row_match = RowMatch.SOURCE_DUPLICATE
        elif not key_confirmed:
            row_match = RowMatch.KEY_UNCONFIRMED
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
        if cell.a1 in self.decimal_checks:
            checks.append(self.decimal_checks[cell.a1])
        total_check = self.total_columns.get(column.letter) or self.total_all
        if total_check:
            checks.append(total_check)

        unit_cell = self._field(reading.record, self.spec.unit_field)
        source_unit = unit_cell.text.strip() if unit_cell else ""
        mapping, match = self._grade_mapping(column, reading.match, key, duplicate, source_unit or None, all(r.key_confirmed for r in group))
        if column.field in self.spec.user_chosen_fields and match is not ColumnMatch.PROFILE:
            mapping = Graded(mapping.color, (*mapping.reasons, Reason("MAP_COLUMN_USER_CHOSEN", "Cột do người dùng chọn ở màn xác nhận")))
        header_unit, form_units, mismatch = self._unit_facts(column, source_unit)
        if not mismatch and (limit is None or not limit.unit) and not source_unit and bool(header_unit) != bool(form_units):
            side = f"nguồn ghi {header_unit}, biểu mẫu không ghi" if header_unit else f"biểu mẫu ghi {form_units[0]}, nguồn không ghi"
            checks.append(check_unit_one_side(f"Chỉ một bên có đơn vị: {side}. Chưa so được, tối đa vàng."))
        confidence = combine(read, mapping, checks)
        if mismatch:
            confidence = _force_red(confidence, Reason("READ_UNIT_MISMATCH", f"{mismatch} Không tự đổi."))
        raw = cell.formula if cell.kind == "empty" else cell.text
        value = Value(type="number", normalized=normalized or raw or "", raw_text=raw or "", unit=source_unit or header_unit)
        self._append(key, column, value, group, confidence, match)

    def _unit_facts(self, column: FormColumn, source_unit: str) -> tuple[str | None, list[str], str | None]:
        """Đơn vị từ cột đơn vị, tiêu đề nguồn, giới hạn, tiêu đề biểu mẫu và tên khai báo. Trả về đơn vị ở tiêu đề nguồn, đơn vị của biểu mẫu, và lời báo nếu lệch (R9b)."""
        resolved = self.plan.columns[column.field]
        header_unit = split_header_unit(resolved.header)[1]
        declared = self.spec.source_fields[column.field].names if column.field in self.spec.source_fields else ()
        form_units = [u for text in column.col_path if (u := split_header_unit(str(text))[1])]
        if column.limit and column.limit.unit:
            form_units.append(column.limit.unit)
        # Tên khai báo chỉ dùng để so: với mô tả tự đoán, đó là tiêu đề nguồn lúc đoán, không phải đơn vị ghi trên biểu mẫu.
        targets = [*form_units, *(u for text in declared if (u := split_header_unit(str(text))[1]))]
        sources = [u for u in (source_unit, header_unit) if u]
        for s in sources:
            for t in targets:
                if not same_unit(s, t):
                    return header_unit, form_units, f"Đơn vị nguồn {s}, biểu mẫu {t}."
        if len(sources) == 2 and not same_unit(*sources):
            return header_unit, form_units, f"Nguồn ghi hai đơn vị: {sources[0]} và {sources[1]}."
        if header_unit and form_units and self.kind in (SourceKind.IMAGE, SourceKind.PDF_IMAGE):
            # Đơn vị đọc từ ảnh cũng phải qua hai lần đọc, mới được coi là khớp.
            sheet = next(s for s in self.source.sheets if s.name == self.plan.sheet)
            cell = sheet.rows[self.plan.header_row][self.idx[column.field]]
            second = split_header_unit(cell.read.second_text)[1] if cell.read and cell.read.second_text else None
            if second is None or not same_unit(second, header_unit):
                return header_unit, form_units, f"Hai lần đọc đơn vị khác nhau: {header_unit} và {second or 'không đọc được'}."
        return header_unit, form_units, None

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
        mapping, match = self._grade_mapping(column, reading.match, key, duplicate, None, all(r.key_confirmed for r in group))
        facts = self._text_read(cell)
        confidence = combine(grade_read(facts), mapping, checks)
        if text.startswith(_FORMULA_PREFIXES):
            # Ô chữ của nguồn không được thành công thức ở biểu mẫu.
            confidence = _force_red(confidence, Reason("READ_TEXT_LOOKS_LIKE_FORMULA", "Chữ trong nguồn bắt đầu như công thức. Backend phải ghi dạng chữ."))
        self._append(key, column, Value(type="text", normalized=text, raw_text=cell.text or text), group, confidence, match)

    def _text_read(self, cell: SourceCell) -> ReadFacts:
        """Ô chữ cũng theo nguồn thật: chữ đọc từ ảnh phải qua hai lần đọc và không bao giờ xanh (Q4)."""
        if self.kind is SourceKind.XLSX:
            return ReadFacts(SourceKind.XLSX, is_formula=cell.formula is not None, has_cached_value=not cell.is_empty)
        if self.kind is SourceKind.PDF_TEXT:
            return ReadFacts(SourceKind.PDF_TEXT, pdf_text_verified=self.source.pdf_text_verified)
        read = cell.read
        readers = tuple(r for r in (read.reader if read else None, read.second_reader if read else None) if r)
        second = read.second_text if read else None
        agree = None if second is None else normalize_text(cell.text).replace(" ", "") == normalize_text(second).replace(" ", "")
        return ReadFacts(
            self.kind,
            low_quality=self.source.low_quality,
            handwriting_declared=self.source.handwriting_declared,
            scanned_text_layer=self.source.scanned_text_layer,
            readers=readers,
            reads_agree=agree,
        )

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
    experimental_image_sources: bool = False,  # Q81: đọc ảnh thành bảng chưa phát hành, chỉ bật khi đo
    experimental_pdf_sources: bool = False,  # Q98: cờ riêng cho PDF có lớp chữ
    previous_month: _Mapping[str, str] | None = None,  # Q119: ô -> số đã duyệt của tháng trước, cùng biểu mẫu
) -> ProposalBatch:
    """Thiếu cột bắt buộc thì ném `MissingFields`, chưa có đề xuất nào được tạo.

    Cùng đầu vào cho cùng đề xuất, trừ `proposal_id` và `created_at`.
    """
    if source.kind == "pdf_text":
        if not experimental_pdf_sources:
            raise FeatureNotReleased("Điền từ PDF")
    elif source.kind != "xlsx" and not experimental_image_sources:
        raise FeatureNotReleased("Điền từ ảnh")
    return _Builder(form, source, apply_header_map(plan, source, header_map, llm_calls), measure_relax_units, previous_month).build()
