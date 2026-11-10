"""Tính màu tin cậy theo docs/ai/confidence_rules.md. Không gọi LLM, không đọc file."""
from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass
from enum import Enum

RULE_VERSION = "0.9"


class Color(str, Enum):
    RED = "red"
    YELLOW = "yellow"
    GREEN = "green"

    @property
    def rank(self) -> int:
        return {"red": 0, "yellow": 1, "green": 2}[self.value]


def lowest(colors: Iterable[Color]) -> Color:
    return min(colors, key=lambda c: c.rank)


class SourceKind(str, Enum):
    XLSX = "xlsx"
    PDF_TEXT = "pdf_text"
    PDF_IMAGE = "pdf_image"
    IMAGE = "image"


class ColumnMatch(str, Enum):
    EXACT = "exact"
    PROFILE = "profile"
    PROFILE_UNCHECKED = "profile_unchecked"
    PROFILE_USER_ANSWER = "profile_user_answer"
    PROFILE_SIMILAR_COLUMNS = "profile_similar_columns"
    NORMALIZED = "normalized"
    INFERRED = "inferred"
    ABBREVIATION = "abbreviation"
    FUZZY = "fuzzy"
    LLM_INFERRED = "llm_inferred"
    AMBIGUOUS = "ambiguous"
    NONE = "none"


class RowMatch(str, Enum):
    KEY_EXACT = "key_exact"
    POSITIONAL = "positional"
    DUPLICATE = "duplicate"
    SOURCE_DUPLICATE = "source_duplicate"
    NOT_FOUND = "not_found"
    KEY_UNCONFIRMED = "key_unconfirmed"


@dataclass(frozen=True)
class Reason:
    code: str
    detail: str


@dataclass(frozen=True)
class Graded:
    color: Color
    reasons: tuple[Reason, ...]


@dataclass(frozen=True)
class ReadFacts:
    """Sự kiện về việc đọc một ô. Điểm tin cậy của OCR cố ý không có ở đây."""

    kind: SourceKind | None
    number_ambiguous: bool = False
    not_a_number: bool = False
    # Excel
    is_formula: bool = False
    has_cached_value: bool = True
    # Ảnh
    handwriting: bool = False
    handwriting_declared: bool = False
    low_quality: bool = False
    readers: tuple[str, ...] = ()
    reads_agree: bool | None = None
    scanned_text_layer: bool = False
    pdf_text_verified: bool = False  # Q98: lớp chữ đã so với OCR đọc từ ảnh dựng của trang. Chưa có bước này thì không xanh


def _engine(reader: str) -> str:
    return reader.split("@", 1)[0]


def grade_read(f: ReadFacts) -> Graded:
    def red(code: str, detail: str) -> Graded:
        return Graded(Color.RED, (Reason(code, detail),))

    if f.kind is None:
        return red("NO_SOURCE", "Không có nguồn")
    if f.kind is SourceKind.XLSX and f.is_formula and not f.has_cached_value:
        return red("READ_FORMULA_NO_CACHED_VALUE", "Ô nguồn có công thức nhưng chưa có giá trị đã lưu")
    if f.number_ambiguous:
        return red("READ_NUMBER_AMBIGUOUS", "Không xác định được dấu thập phân")
    if f.not_a_number:
        return red("READ_NOT_A_NUMBER", "Chữ đọc được không phải là số")

    if f.kind in (SourceKind.IMAGE, SourceKind.PDF_IMAGE):
        if f.handwriting_declared:
            return red("READ_HANDWRITING_DECLARED", "Người dùng khai báo ảnh có chữ viết tay")
        if f.handwriting:
            return red("READ_HANDWRITING", "Chữ viết tay")
        if f.low_quality:
            return red("READ_IMAGE_LOW_QUALITY", "Ảnh dưới ngưỡng chất lượng")
        if len({_engine(r) for r in f.readers}) < 2 or f.reads_agree is None:
            return red("READ_IMAGE_SINGLE_READ", "Chưa có hai lần đọc độc lập")
        if not f.reads_agree:
            return red("READ_IMAGE_TWO_READS_DISAGREE", "Hai lần đọc độc lập khác nhau")
        reasons = [Reason("READ_IMAGE_TWO_READS_AGREE", "Hai lần đọc độc lập cho cùng một số")]
        if f.scanned_text_layer:
            reasons.append(Reason("READ_PDF_SCANNED_TEXT_LAYER", "Lớp chữ nằm trên ảnh quét, xử lý như ảnh"))
        return Graded(Color.YELLOW, tuple(reasons))

    if f.kind is SourceKind.XLSX:
        if f.is_formula:
            return Graded(Color.YELLOW, (Reason("READ_XLSX_CACHED_FORMULA", "Dùng giá trị đã lưu của công thức nguồn"),))
        return Graded(Color.GREEN, (Reason("READ_XLSX_DIRECT", "Đọc trực tiếp từ Excel"),))
    if not f.pdf_text_verified:
        return Graded(Color.YELLOW, (Reason("READ_PDF_TEXT_UNVERIFIED", "Đọc từ lớp chữ PDF, chưa so với thứ hiển thị trên trang"),))
    return Graded(Color.GREEN, (Reason("READ_PDF_TEXT_DIRECT", "Lớp chữ PDF khớp với OCR đọc từ ảnh dựng của trang"),))


_COLUMN_GRADES = {
    ColumnMatch.EXACT: (Color.GREEN, "MAP_COLUMN_EXACT", "Tiêu đề cột khớp"),
    ColumnMatch.NORMALIZED: (Color.GREEN, "MAP_COLUMN_EXACT", "Tiêu đề cột khớp sau khi bỏ dấu, chữ hoa thường"),
    ColumnMatch.PROFILE: (Color.GREEN, "MAP_COLUMN_PROFILE", "Khớp hồ sơ biểu mẫu đã xác nhận"),
    ColumnMatch.PROFILE_UNCHECKED: (Color.YELLOW, "MAP_PROFILE_NO_INDEPENDENT_CHECK", "Khớp hồ sơ nhưng cột không có kiểm tra độc lập"),
    # Q114, Q115: ghép nhờ người dùng chọn cột hoặc trả lời "Cùng một thứ". Bấm nhầm thì giới hạn ±3SD không bắt được hết.
    ColumnMatch.PROFILE_USER_ANSWER: (Color.YELLOW, "MAP_PROFILE_USER_ANSWER", "Khớp hồ sơ, nhưng cách ghép cột này do người dùng chọn hoặc trả lời"),
    ColumnMatch.PROFILE_SIMILAR_COLUMNS: (Color.YELLOW, "MAP_PROFILE_SIMILAR_COLUMNS", "Khớp hồ sơ, nhưng nguồn có nhiều cột số có thể là kết quả"),
    ColumnMatch.INFERRED: (Color.YELLOW, "MAP_COLUMN_INFERRED", "Máy suy ra, chưa có hồ sơ được xác nhận"),
    ColumnMatch.ABBREVIATION: (Color.YELLOW, "MAP_COLUMN_ABBREVIATION", "Khớp qua từ điển viết tắt"),
    ColumnMatch.FUZZY: (Color.YELLOW, "MAP_COLUMN_FUZZY", "Khớp gần đúng"),
    ColumnMatch.LLM_INFERRED: (Color.YELLOW, "MAP_COLUMN_LLM", "LLM suy luận cột"),
    ColumnMatch.AMBIGUOUS: (Color.RED, "MAP_COLUMN_AMBIGUOUS", "Nhiều cột có thể khớp"),
    ColumnMatch.NONE: (Color.RED, "MAP_COLUMN_NONE", "Không tìm thấy cột"),
}
_ROW_GRADES = {
    RowMatch.KEY_EXACT: (Color.GREEN, "MAP_ROW_KEY_EXACT", "Khóa dòng khớp đúng một dòng"),
    RowMatch.POSITIONAL: (Color.YELLOW, "MAP_ROW_POSITIONAL", "Không có cột khóa, dùng vị trí dòng"),
    RowMatch.DUPLICATE: (Color.RED, "MAP_ROW_DUPLICATE", "Khóa cho nhiều dòng"),
    RowMatch.SOURCE_DUPLICATE: (Color.RED, "SOURCE_DUPLICATE_RESULT", "Nguồn có nhiều kết quả cùng khóa"),
    RowMatch.NOT_FOUND: (Color.RED, "MAP_ROW_NOT_FOUND", "Khóa không tìm thấy dòng nào"),
    RowMatch.KEY_UNCONFIRMED: (Color.RED, "MAP_ROW_KEY_UNCONFIRMED", "Khóa dòng từ ảnh chưa được hai lần đọc xác nhận"),
}


def grade_map(column: ColumnMatch, row: RowMatch) -> Graded:
    col_color, col_code, col_detail = _COLUMN_GRADES[column]
    row_color, row_code, row_detail = _ROW_GRADES[row]
    return Graded(
        lowest((col_color, row_color)),
        (Reason(col_code, col_detail), Reason(row_code, row_detail)),
    )


@dataclass(frozen=True)
class CheckResult:
    """Một kiểm chéo không đạt. `ceiling` là màu cao nhất ô còn được giữ."""

    ceiling: Color
    reason: Reason


def check_implausible(detail: str) -> CheckResult:
    return CheckResult(Color.RED, Reason("CHECK_IMPLAUSIBLE", detail))


def check_out_of_3sd(detail: str) -> CheckResult:
    return CheckResult(Color.YELLOW, Reason("CHECK_OUT_OF_3SD", detail))


def check_printed_total(detail: str) -> CheckResult:
    return CheckResult(Color.RED, Reason("CHECK_PRINTED_TOTAL_MISMATCH", detail))


def check_total_not_comparable(detail: str) -> CheckResult:
    return CheckResult(Color.RED, Reason("CHECK_TOTAL_NOT_COMPARABLE", detail))


def check_lot_mismatch(detail: str) -> CheckResult:
    return CheckResult(Color.RED, Reason("CHECK_LOT_MISMATCH", detail))


def check_column_suspect(detail: str) -> CheckResult:
    return CheckResult(Color.RED, Reason("CHECK_COLUMN_SUSPECT", detail))


def check_not_in_list(detail: str) -> CheckResult:
    return CheckResult(Color.RED, Reason("CHECK_NOT_IN_LIST", detail))


def check_decimal_places(detail: str) -> CheckResult:
    return CheckResult(Color.YELLOW, Reason("CHECK_DECIMAL_PLACES", detail))


def check_unit_one_side(detail: str) -> CheckResult:
    return CheckResult(Color.YELLOW, Reason("CHECK_UNIT_ONE_SIDE", detail))


def check_decimal_shift(detail: str) -> CheckResult:
    return CheckResult(Color.RED, Reason("CHECK_DECIMAL_SHIFT", detail))


@dataclass(frozen=True)
class Confidence:
    read_color: Color
    map_color: Color
    final_color: Color
    reasons: tuple[Reason, ...] = ()


def combine(read: Graded, mapping: Graded, checks: Iterable[CheckResult] = ()) -> Confidence:
    """final = min(đọc, ghép), rồi kiểm chéo chỉ có thể hạ."""
    final = lowest((read.color, mapping.color))
    reasons = [*read.reasons, *mapping.reasons]
    for check in checks:
        final = lowest((final, check.ceiling))
        reasons.append(check.reason)
    return Confidence(read.color, mapping.color, final, tuple(reasons))


def grade_formula(
    input_colors: Iterable[Color | None],
    proposed_by_llm: bool,
    input_unknown: bool = False,
    *,
    empty_inputs: int = 0,
    skipped_inputs: int = 0,
    pending_inputs: int = 0,
    total_inputs: int = 0,
) -> tuple[Color | None, tuple[Reason, ...]]:
    """Màu ô công thức. None trong đầu vào là ô nhập tay và không tham gia.

    Đề xuất còn chờ duyệt: đưa màu của nó vào `input_colors`, và đếm vào `pending_inputs`, vì số chưa vào lưới nên công thức đang tính thiếu ô đó.
    `skipped_inputs`: ô có đề xuất mà người dùng bỏ qua.
    `empty_inputs`: ô trống không có đề xuất. Công thức vẫn ra số trên ít ô hơn nên không được xanh, và lý do phải nói số ô.
    """
    if input_unknown:
        return Color.RED, (Reason("FORMULA_INPUT_UNKNOWN", "Không xác định được ô đầu vào"),)
    colors = [c for c in input_colors if c is not None]
    reasons: tuple[Reason, ...] = ()
    if proposed_by_llm:
        colors.append(Color.YELLOW)
        reasons = (Reason("FORMULA_LLM_PROPOSED", "Công thức do LLM đề xuất"),)
    missing = empty_inputs + skipped_inputs + pending_inputs
    used = f"{total_inputs - missing}/{total_inputs}" if total_inputs else "ít"
    if pending_inputs:
        colors.append(Color.YELLOW)
        reasons = (*reasons, Reason("FORMULA_INPUT_PENDING", f"Công thức chưa tính {pending_inputs} ô còn chờ duyệt, đang tính trên {used} ô"))
    if skipped_inputs:
        colors.append(Color.YELLOW)
        reasons = (*reasons, Reason("FORMULA_INPUT_SKIPPED", f"Công thức đang tính thiếu {skipped_inputs} ô đã bỏ qua, trên {used} ô"))
    if empty_inputs and colors:  # chỉ có ô nhập tay thì AI không tham gia, vẫn không có màu
        colors.append(Color.YELLOW)
        reasons = (*reasons, Reason("FORMULA_INPUT_EMPTY", f"Công thức đang tính trên {used} ô, {empty_inputs} ô còn trống"))
    if not colors:
        return None, reasons  # chỉ có ô nhập tay: không có màu, không phải xanh
    return lowest(colors), reasons
