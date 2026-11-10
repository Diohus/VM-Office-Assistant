"""Quy tắc màu phải được nối vào đường chạy thật, không chỉ có hàm (Q95). Token cấy, không OCR, không chứng minh gì với ảnh thật."""
from __future__ import annotations

import copy
import datetime as dt
import re
from collections import Counter, defaultdict
from pathlib import Path
from types import SimpleNamespace

import openpyxl
import doc_snapshot
import pytest
from test_images import _run, _source, _template, _tokens

import vmai
from vmai import MissingFields, infer_spec, load_spec, plan_headers, propose, read_form, read_source
from vmai.confidence import Color, ReadFacts, RowMatch, SourceKind, grade_read
from vmai.mapping import match_field, same_unit, split_header_unit
from vmai.propose import _Builder
from vmai.source_xlsx import CellRead, SourceCell
from vmai.table_read import source_from_tokens

SRC = Path(vmai.__file__).parent
FULL_TREE = (SRC.parents[2] / "project-memory").is_dir()  # kho công khai không có project-memory/ và docs/
DOC = SRC.parents[2] / "docs" / "ai" / "confidence_rules.md"

# Mã có trong tài liệu nhưng cố ý chưa phát ra ở `vmai`. Thêm vào đây phải có lý do, và phải ghi trong báo cáo đối chiếu.
NOT_IN_VMAI = {
    "STALE_SOURCE_CHANGED": "chưa làm ở đâu, Dev phần 3",
    "STALE_VALUE_CHANGED": "chưa làm ở đâu, Dev phần 3",
    "STALE_FORMULA_CHANGED": "chưa làm ở đâu, Dev phần 3",
    "STALE_CELL_NOT_FOUND": "chưa làm ở đâu, Dev phần 3",
    "STALE_POSITIONAL_LINK": "chưa làm ở đâu, Dev phần 3",
    "STALE_PROFILE_REVOKED": "chưa làm ở đâu, Dev phần 3",
}
# Đường trong bảng màu chưa có nơi nào đi vào. Xem bảng đối chiếu trong docs/ai/rule_audit.md.
UNREACHED_ROW_MATCHES = {RowMatch.POSITIONAL, RowMatch.NOT_FOUND}


def _source_text(skip_confidence: bool = True) -> str:
    return "\n".join(p.read_text(encoding="utf-8") for p in SRC.glob("*.py") if not (skip_confidence and p.name == "confidence.py"))


def test_every_check_function_has_a_caller_outside_its_definition():
    definitions = re.findall(r"def (check_\w+)", (SRC / "confidence.py").read_text(encoding="utf-8"))
    others = _source_text()
    assert definitions
    missing = [fn for fn in definitions if not re.search(rf"\b{fn}\(", others)]
    assert not missing, f"Quy tắc có hàm nhưng không được gọi: {missing}"


def test_every_row_match_is_reached_or_listed():
    others = _source_text()
    for member in RowMatch:
        reached = f"RowMatch.{member.name}" in others
        assert reached is (member not in UNREACHED_ROW_MATCHES), f"{member}: đổi danh sách và bảng đối chiếu"


def test_every_code_in_the_rules_doc_exists_in_code_or_is_listed():
    codes = set(doc_snapshot.load()["confidence_rules_codes"])  # bản chụp từ docs/ai/confidence_rules.md
    text = _source_text(skip_confidence=False)
    missing = {c for c in codes if c not in text} - set(NOT_IN_VMAI)
    assert not missing, f"Tài liệu có mã mà code không có: {missing}"
    for code in NOT_IN_VMAI:
        assert code not in text, f"{code} đã có trong vmai, bỏ khỏi danh sách"


# ---- Tổng in trên phiếu


def _footer(line, y=400):
    return [{"text": word, "bbox": [50 + i * 80, y, 120 + i * 80, y + 20]} for i, word in enumerate(line.split())]


def _colors(batch):
    return Counter(c.confidence.final_color.value for c in batch.changes)


def _codes(change):
    return {r.code for r in change.confidence.reasons}


@pytest.mark.parametrize("line", ["Tổng số mẫu: 17", "TONG SO MAU 17", "Tổng: 17", "Tổng cộng: 17"])
def test_a08_printed_count_mismatch_makes_every_number_red(tmp_path, line):
    first = [*_tokens(), *_footer(line)]
    _, batch = _run(tmp_path, first, copy.deepcopy(first))
    assert _colors(batch) == {"red": 10}
    assert all("CHECK_PRINTED_TOTAL_MISMATCH" in _codes(c) for c in batch.changes)
    assert "CHECK_PRINTED_TOTAL_MISMATCH" in {n.reason for n in batch.notices}


@pytest.mark.parametrize("line", ["Tổng số mẫu: 10", "Tổng số mẫu: 5"])
def test_printed_count_that_matches_rows_or_results_does_not_lower_or_raise(tmp_path, line):
    first = [*_tokens(), *_footer(line)]
    _, batch = _run(tmp_path, first, copy.deepcopy(first))
    assert _colors(batch) == {"yellow": 10}


def test_printed_count_misread_by_second_engine_is_not_trusted(tmp_path):
    first, second = [*_tokens(), *_footer("Tổng số mẫu: 10")], [*_tokens(), *_footer("Tổng số mẫu: 18")]
    _, batch = _run(tmp_path, first, second)
    assert _colors(batch) == {"red": 10}


def test_a_title_with_a_number_is_not_a_total(tmp_path):
    first = [*_tokens(), *_footer("Tổng hợp tháng 9")]
    _, batch = _run(tmp_path, first, copy.deepcopy(first))
    assert _colors(batch) == {"yellow": 10}


def _excel(tmp_path, template_rows, source_rows, name="s.xlsx"):
    template = tmp_path / "t.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "[MẪU GIẢ]"
    for r, row in enumerate(template_rows, start=3):
        for c, value in enumerate(row, start=1):
            if value is not None:
                ws.cell(r, c, value)
    wb.save(template)
    source = tmp_path / name
    wb = openpyxl.Workbook()
    for row in source_rows:
        wb.active.append(row)
    wb.save(source)
    return template, source


def _excel_batch(template, source):
    facts = read_source(source)
    guess = infer_spec(read_source(template), facts)
    assert guess.spec is not None, [q.code for q in guess.questions]
    form = read_form(template, guess.spec)
    return propose(form, facts, plan_headers(form, facts))


DAYS = [dt.datetime(2026, 9, d) for d in range(1, 7)]


def test_excel_total_row_mismatch_lowers_only_that_column(tmp_path):
    template, source = _excel(
        tmp_path,
        [["Ngày", "Natri", "Kali"], *[[d] for d in DAYS]],
        [["Ngày", "Natri", "Kali"], *[[d, 140, 4.5] for d in DAYS], ["Tổng", 999, 27.0]],
    )
    batch = _excel_batch(template, source)
    by_col = {c.target.a1_at_write[0]: c.confidence.final_color for c in batch.changes}
    assert by_col["B"] is Color.RED and by_col["C"] is not Color.RED


def test_excel_total_row_that_matches_changes_nothing(tmp_path):
    template, source = _excel(
        tmp_path,
        [["Ngày", "Natri", "Kali"], *[[d] for d in DAYS]],
        [["Ngày", "Natri", "Kali"], *[[d, 140, 4.5] for d in DAYS], ["Tổng", 840, 27.0]],
    )
    assert not any("CHECK_PRINTED_TOTAL_MISMATCH" in _codes(c) for c in _excel_batch(template, source).changes)


# ---- Hai engine cùng mất dấu thập phân


def _agree(values):
    tokens = _tokens(values=values)
    return tokens, copy.deepcopy(tokens)


def _shift_batch(tmp_path, values):
    first, second = _agree(values)
    return _run(tmp_path, first, second)[1]


BASE = {r: ("5,10", "16,20") for r in range(5)}


@pytest.mark.parametrize(
    "bad",
    ["510", "51,00", "0,51"],
    ids=["mat_dau", "dich_dau_cung_so_chu_so", "dich_dau_ve_trai"],
)
def test_both_engines_agreeing_on_a_shifted_decimal_is_red_not_yellow(tmp_path, bad):
    batch = _shift_batch(tmp_path, {**BASE, 2: (bad, "16,20")})
    cell = next(c for c in batch.changes if c.target.a1_at_write == "B7")
    assert cell.confidence.final_color is Color.RED and "CHECK_DECIMAL_SHIFT" in _codes(cell)
    assert _colors(batch)["yellow"] == 9


def test_most_of_a_column_lost_the_decimal_then_the_whole_column_is_red(tmp_path):
    batch = _shift_batch(tmp_path, {r: ("510" if r != 4 else "5,10", "16,20") for r in range(5)})
    assert all(c.confidence.final_color is Color.RED for c in batch.changes if c.target.a1_at_write.startswith("B"))


def test_known_gap_whole_column_lost_the_decimal_stays_yellow(tmp_path):
    """Rủi ro còn lại, đã báo PM: cả cột cùng mất dấu thì không còn ô nào làm mốc. Đổi hành vi thì sửa test và báo cáo."""
    batch = _shift_batch(tmp_path, {r: ("510", "16,20") for r in range(5)})
    assert all(c.confidence.final_color is Color.YELLOW for c in batch.changes if c.target.a1_at_write.startswith("B"))


def test_wide_integer_column_is_not_flagged_as_a_shift(tmp_path):
    batch = _shift_batch(tmp_path, {0: ("3", "16,20"), 1: ("48", "16,20"), 2: ("120", "16,20"), 3: ("260", "16,20"), 4: ("480", "16,20")})
    assert not any("CHECK_DECIMAL_SHIFT" in _codes(c) for c in batch.changes)


def test_cells_the_engines_disagree_on_are_not_used_as_the_reference(tmp_path):
    first = _tokens(values={**{r: ("44", "16,20") for r in range(4)}, 4: ("4,4", "16,20")})
    second = _tokens(values={r: ("4,4", "16,20") for r in range(5)})
    _, batch = _run(tmp_path, first, second)
    last = next(c for c in batch.changes if c.target.a1_at_write == "B9")
    assert last.confidence.final_color is Color.YELLOW


# ---- Số chữ số thập phân


def test_decimal_places_differ_caps_excel_text_number_at_yellow(tmp_path):
    template, source = _excel(
        tmp_path,
        [["Ngày", "Natri"], *[[d] for d in DAYS]],
        [["Ngày", "Natri"], *[[d, text] for d, text in zip(DAYS, ["5,42", "5,31", "5,50", "5,47", "5,12", "5,4"], strict=True)]],
    )
    batch = _excel_batch(template, source)
    odd = next(c for c in batch.changes if c.value.raw_text == "5,4")
    assert "CHECK_DECIMAL_PLACES" in _codes(odd) and odd.confidence.final_color is Color.YELLOW
    assert sum("CHECK_DECIMAL_PLACES" in _codes(c) for c in batch.changes) == 1


def test_real_excel_numbers_are_not_checked_for_decimal_places(tmp_path):
    template, source = _excel(tmp_path, [["Ngày", "Natri"], *[[d] for d in DAYS]], [["Ngày", "Natri"], *[[d, v] for d, v in zip(DAYS, [5.42, 5.31, 5.5, 5.47, 5.12, 5.0], strict=True)]])
    assert not any("CHECK_DECIMAL_PLACES" in _codes(c) for c in _excel_batch(template, source).changes)


# ---- Khóa dòng từ ảnh


def test_key_read_differently_by_the_second_engine_is_red(tmp_path):
    first, second = _tokens(), _tokens()
    for token in first:
        if token["text"] == "05/09/2026":
            token["text"] = "03/09/2026"
    _, batch = _run(tmp_path, first, second)
    unconfirmed = [c for c in batch.changes if "MAP_ROW_KEY_UNCONFIRMED" in _codes(c)]
    assert len(unconfirmed) == 2 and all(c.confidence.final_color is Color.RED and c.confidence.map_color is Color.RED for c in unconfirmed)


def test_key_with_no_second_read_is_red(tmp_path):
    first, second = _tokens(), _tokens()
    second = [t for t in second if t["text"] != "04/09/2026"]
    _, batch = _run(tmp_path, first, second)
    assert sum("MAP_ROW_KEY_UNCONFIRMED" in _codes(c) for c in batch.changes) == 2


# ---- Viết tay, lớp chữ quét, ô chữ từ ảnh


def _batch_from(tmp_path, source):
    template = _template(tmp_path / "t.xlsx")
    guess = infer_spec(read_source(template), source, frozenset({"LEVEL_NAME_DIFFERS"}), experimental_image_sources=True)
    form = read_form(template, guess.spec)
    return propose(form, source, plan_headers(form, source), experimental_image_sources=True)


def test_r5b_declared_handwriting_is_red(tmp_path):
    tokens = _tokens()
    source = source_from_tokens("a.jpg", "x", [(1, tokens, copy.deepcopy(tokens))], kind="image", first_reader="ocr:a@1", second_reader="ocr:b@1", handwriting_declared=True)
    batch = _batch_from(tmp_path, source)
    assert _colors(batch) == {"red": 10} and all("READ_HANDWRITING_DECLARED" in _codes(c) for c in batch.changes)


def test_r11_scanned_text_layer_is_treated_as_an_image_and_says_so(tmp_path):
    tokens = _tokens()
    source = source_from_tokens("a.pdf", "x", [(1, tokens, copy.deepcopy(tokens))], kind="pdf_image", first_reader="ocr:a@1", second_reader="ocr:b@1", scanned_text_layer=True)
    batch = _batch_from(tmp_path, source)
    assert _colors(batch) == {"yellow": 10} and all("READ_PDF_SCANNED_TEXT_LAYER" in _codes(c) for c in batch.changes)


def _text_facts(kind, second):
    cell = SourceCell(a1="A1", kind="text", value="Đạt", read=CellRead(page=1, bbox=(0, 0, 1, 1), reader="ocr:a@1", second_reader="ocr:b@1", second_text=second))
    return _Builder._text_read(SimpleNamespace(kind=kind, source=SimpleNamespace(low_quality=False, handwriting_declared=False, scanned_text_layer=False)), cell)


def test_text_cell_from_an_image_is_never_green():
    assert grade_read(_text_facts(SourceKind.IMAGE, "Đạt")).color is Color.YELLOW
    assert grade_read(_text_facts(SourceKind.IMAGE, "Dat")).color is Color.YELLOW  # bỏ dấu vẫn là một chữ
    assert grade_read(_text_facts(SourceKind.IMAGE, "Đạo")).color is Color.RED
    assert grade_read(_text_facts(SourceKind.IMAGE, None)).color is Color.RED
    assert grade_read(ReadFacts(SourceKind.XLSX)).color is Color.GREEN


# ---- Đơn vị trong tiêu đề (R9b). Bỏ đơn vị để ghép là cách dễ nhất và nguy hiểm nhất.


@pytest.mark.parametrize(
    ("text", "name", "unit"),
    [
        ("HGB (g/L)", "HGB", "g/L"),
        ("HGB [g/dL]", "HGB", "g/dL"),
        ("Tỷ lệ (%)", "Tỷ lệ", "%"),
        ("Glucose (máy 2)", "Glucose (máy 2)", None),
        ("Mức (1)", "Mức (1)", None),
        ("Tỷ số (1/2)", "Tỷ số (1/2)", None),
        ("HGB", "HGB", None),
    ],
)
def test_split_header_unit(text, name, unit):
    assert split_header_unit(text) == (name, unit)


@pytest.mark.parametrize(("a", "b", "same"), [("g/L", "g/L", True), ("g/dL", "g/dl", True), ("µmol/L", "umol/L", True), ("g/L", "g/dL", False), ("G/L", "g/L", False), ("mmol/L", "umol/L", False), ("mg/dL", "mg/L", False)])
def test_same_unit(a, b, same):
    assert same_unit(a, b) is same


def _unit_case(tmp_path, template_header, source_header, value=13.5):
    return _excel(
        tmp_path,
        [["Ngày", template_header], *[[d] for d in DAYS]],
        [["Ngày", source_header], *[[d, value] for d in DAYS]],
    )


@pytest.mark.parametrize(
    ("template_header", "source_header"),
    [("HGB (g/L)", "HGB (g/dL)"), ("HGB (g/dL)", "HGB (g/L)"), ("WBC (G/L)", "WBC (g/L)"), ("Kết quả (mmol/L)", "Kết quả (umol/L)"), ("Hgb (G/L)", "HGB [g/L]")],
)
def test_trap_same_name_other_unit_is_never_matched_silently(tmp_path, template_header, source_header):
    template, source = _unit_case(tmp_path, template_header, source_header)
    batch = _excel_batch(template, source)
    assert batch.changes
    assert all(c.confidence.final_color is Color.RED and "READ_UNIT_MISMATCH" in _codes(c) for c in batch.changes)
    assert "READ_UNIT_MISMATCH" in {n.reason for n in batch.notices}
    # Không tự đổi: số giữ nguyên như nguồn.
    assert {c.value.normalized for c in batch.changes} == {"13.5"}


def test_trap_other_unit_with_a_declared_form_does_not_reach_yellow(tmp_path):
    """Hồ sơ hoặc mô tả đã khai `HGB (g/L)`, nguồn lần sau đổi sang g/dL. Tiêu đề gần giống vẫn không được xanh hay vàng."""
    template, source = _unit_case(tmp_path, "HGB (g/L)", "HGB (g/L)")
    guess = infer_spec(read_source(template), read_source(source))
    form = read_form(template, guess.spec)
    _, other = _unit_case(tmp_path, "HGB (g/L)", "HGB (g/dL)")
    facts = read_source(other)
    try:
        batch = propose(form, facts, plan_headers(form, facts))
    except MissingFields:
        return  # không ghép được thì hỏi, cũng là không im lặng
    assert all(c.confidence.final_color is Color.RED for c in batch.changes)


def test_trap_on_image_other_unit_is_red(tmp_path):
    template = tmp_path / "t.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A3"], ws["B3"], ws["C3"] = "Ngày", "HGB (g/L)", "PLT"
    for i in range(5):
        ws[f"A{4 + i}"] = dt.date(2026, 9, 1 + i)
    wb.save(template)
    tokens = _tokens(headers=["Ngày", "HGB (g/dL)", "PLT"])
    source = _source(tokens, copy.deepcopy(tokens))
    guess = infer_spec(read_source(template), source, frozenset(), experimental_image_sources=True)
    assert guess.spec is not None, [q.code for q in guess.questions]
    form = read_form(template, guess.spec)
    batch = propose(form, source, plan_headers(form, source), experimental_image_sources=True)
    hgb = [c for c in batch.changes if c.target.a1_at_write.startswith("B")]
    assert hgb and all(c.confidence.final_color is Color.RED and "READ_UNIT_MISMATCH" in _codes(c) for c in hgb)


def test_on_image_unit_read_differently_by_the_second_engine_is_red(tmp_path):
    template = tmp_path / "t.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A3"], ws["B3"], ws["C3"] = "Ngày", "HGB (g/L)", "PLT"
    for i in range(5):
        ws[f"A{4 + i}"] = dt.date(2026, 9, 1 + i)
    wb.save(template)
    first = _tokens(headers=["Ngày", "HGB (g/L)", "PLT"])
    second = _tokens(headers=["Ngày", "HGB (g/dL)", "PLT"])
    source = _source(first, second)
    guess = infer_spec(read_source(template), source, frozenset(), experimental_image_sources=True)
    form = read_form(template, guess.spec)
    batch = propose(form, source, plan_headers(form, source), experimental_image_sources=True)
    assert all(c.confidence.final_color is Color.RED for c in batch.changes if c.target.a1_at_write.startswith("B"))


def test_same_unit_written_differently_is_not_a_mismatch(tmp_path):
    template, source = _unit_case(tmp_path, "HGB (g/dL)", "HGB (g/dl)")
    assert not any("READ_UNIT_MISMATCH" in _codes(c) for c in _excel_batch(template, source).changes)


@pytest.mark.parametrize(("template_header", "source_header"), [("HGB", "HGB (g/L)"), ("HGB (g/L)", "HGB")])
def test_unit_on_one_side_only_is_mapped_but_never_green(tmp_path, template_header, source_header):
    template, source = _unit_case(tmp_path, template_header, source_header)
    batch = _excel_batch(template, source)
    assert batch.changes
    assert all(c.confidence.final_color is Color.YELLOW and "CHECK_UNIT_ONE_SIDE" in _codes(c) for c in batch.changes)


def test_source_unit_is_shown_to_the_reviewer(tmp_path):
    template, source = _unit_case(tmp_path, "HGB", "HGB (g/L)")
    assert {c.value.unit for c in _excel_batch(template, source).changes} == {"g/L"}


def test_two_source_columns_same_name_picks_only_the_same_unit(tmp_path):
    template, source = _excel(
        tmp_path,
        [["Ngày", "HGB (g/L)"], *[[d] for d in DAYS]],
        [["Ngày", "HGB (g/dL)", "HGB (g/L)"], *[[d, 13.5, 135] for d in DAYS]],
    )
    batch = _excel_batch(template, source)
    assert {c.value.normalized for c in batch.changes} == {"135"}
    assert not any("READ_UNIT_MISMATCH" in _codes(c) for c in batch.changes)


def test_two_source_columns_and_no_unit_on_the_form_asks(tmp_path):
    template, source = _excel(
        tmp_path,
        [["Ngày", "HGB"], *[[d] for d in DAYS]],
        [["Ngày", "HGB (g/dL)", "HGB (g/L)"], *[[d, 13.5, 135] for d in DAYS]],
    )
    guess = infer_spec(read_source(template), read_source(source))
    assert guess.spec is None and guess.questions


def test_fuzzy_match_never_crosses_units():
    fields = {"hgb": ("HGB (g/L)",)}
    assert match_field("HGB (g/dL)", fields, {}).field is None
    assert match_field("HGBs (g/L)", fields, {}).field == "hgb"


def test_parenthesis_that_is_not_a_unit_still_separates_columns(tmp_path):
    template, source = _unit_case(tmp_path, "Glucose", "Glucose (máy 2)")
    guess = infer_spec(read_source(template), read_source(source))
    assert guess.spec is None and guess.questions


# ---- Đã nối từ trước nhưng chưa test nào đi qua propose (đo bằng cách ghi mã lý do thật sự phát ra, 05/11)


def test_low_quality_image_is_red(tmp_path):
    tokens = _tokens()
    source = source_from_tokens("a.jpg", "x", [(1, tokens, copy.deepcopy(tokens))], kind="image", first_reader="ocr:a@1", second_reader="ocr:b@1", low_quality=True)
    batch = _batch_from(tmp_path, source)
    assert _colors(batch) == {"red": 10} and all("READ_IMAGE_LOW_QUALITY" in _codes(c) for c in batch.changes)


def test_cell_missing_from_the_second_read_is_red(tmp_path):
    first, second = _tokens(), _tokens()
    second = [t for i, t in enumerate(second) if not (t["text"] == "5,10" and i == 4)]
    _, batch = _run(tmp_path, first, second)
    single = [c for c in batch.changes if "READ_IMAGE_SINGLE_READ" in _codes(c)]
    assert len(single) == 1 and single[0].confidence.final_color is Color.RED


def test_same_engine_twice_is_not_two_reads(tmp_path):
    tokens = _tokens()
    source = source_from_tokens("a.jpg", "x", [(1, tokens, copy.deepcopy(tokens))], kind="image", first_reader="ocr:a@1", second_reader="ocr:a@2")
    batch = _batch_from(tmp_path, source)
    assert _colors(batch) == {"red": 10} and all("READ_IMAGE_SINGLE_READ" in _codes(c) for c in batch.changes)


def test_pdf_text_layer_is_at_most_yellow_until_compared_with_ocr(tmp_path):
    tokens = _tokens()
    template = _template(tmp_path / "t.xlsx")
    guess = infer_spec(read_source(template), _source(tokens, copy.deepcopy(tokens)), frozenset({"LEVEL_NAME_DIFFERS"}), experimental_image_sources=True)
    form = read_form(template, guess.spec)
    source = source_from_tokens("a.pdf", "x", [(1, tokens, None)], kind="pdf_text", first_reader="pdf_text", second_reader=None)
    batch = propose(form, source, plan_headers(form, source), experimental_pdf_sources=True)
    assert batch.changes and all("READ_PDF_TEXT_UNVERIFIED" in _codes(c) and c.confidence.read_color is Color.YELLOW for c in batch.changes)


def test_duplicate_key_in_the_form_is_red(tmp_path):
    template, source = _excel(tmp_path, [["Ngày", "Natri"], DAYS[:1], *[[d] for d in DAYS]], [["Ngày", "Natri"], *[[d, 140] for d in DAYS]])
    facts = read_source(source)
    guess = infer_spec(read_source(template), facts, frozenset({"KEY_NOT_UNIQUE"}))
    form = read_form(template, guess.spec)
    batch = propose(form, facts, plan_headers(form, facts))
    duplicated = [c for c in batch.changes if "MAP_ROW_DUPLICATE" in _codes(c)]
    assert duplicated and all(c.confidence.final_color is Color.RED for c in duplicated)


def test_fuzzy_header_is_at_most_yellow(tmp_path):
    template, source = _excel(tmp_path, [["Ngày", "Natri huyết thanh"], *[[d] for d in DAYS]], [["Ngày", "Natri huyết thanh"], *[[d, 140] for d in DAYS]])
    guess = infer_spec(read_source(template), read_source(source))
    form = read_form(template, guess.spec)
    other = tmp_path / "o.xlsx"
    wb = openpyxl.Workbook()
    for row in [["Ngày", "Natri huyet thanhh"], *[[d, 140] for d in DAYS]]:
        wb.active.append(row)
    wb.save(other)
    facts = read_source(other)
    batch = propose(form, facts, plan_headers(form, facts))
    assert batch.changes and all("MAP_COLUMN_FUZZY" in _codes(c) and c.confidence.final_color is not Color.GREEN for c in batch.changes)


def test_value_outside_the_option_list_is_red(tmp_path):
    template, source = _excel(tmp_path, [["Ngày", "Natri"], *[[d] for d in DAYS]], [["Ngày", "Natri", "KL"], *[[d, 140, "Đạt"] for d in DAYS[:-1]], [DAYS[-1], 140, "Lạ"]])
    facts = read_source(source)
    guess = infer_spec(read_source(template), facts)
    wb = openpyxl.load_workbook(template)
    wb.active["C3"] = "Kết luận"
    options = wb.create_sheet("DS")
    options["A1"], options["A2"] = "Đạt", "Không đạt"
    wb.save(template)
    data = guess.spec.model_dump()
    data["source_fields"]["kl"] = {"names": ["KL"], "required": True}
    data["table"]["value_columns"] = [*data["table"]["value_columns"], {"column": "C", "field": "kl", "type": "text", "category": True}]
    data["options"] = [{"applies_to": "C", "sheet": "DS", "column": "A", "first_row": 1}]
    form = read_form(template, load_spec(data))
    batch = propose(form, facts, plan_headers(form, facts))
    text = {c.value.normalized: c for c in batch.changes if c.target.a1_at_write.startswith("C")}
    assert "CHECK_NOT_IN_LIST" in _codes(text["Lạ"]) and text["Lạ"].confidence.final_color is Color.RED
    assert "CHECK_NOT_IN_LIST" not in _codes(text["Đạt"])


# ---- Dòng tổng phải đi tới được bước so, cả với bảng khóa mã chữ. Tách đọc thiếu với tổng lệch (Q97).


CODES = [f"MA{i:03d}" for i in range(1, 9)]


def _code_tokens(values, footer=None):
    xs = [50, 200, 400]
    out = [{"text": h, "bbox": [xs[i], 100, xs[i] + 90, 120]} for i, h in enumerate(["Mã", "Số lượng", "Đơn giá"])]
    for r, code in enumerate(CODES):
        y = 130 + r * 30
        out.append({"text": code, "bbox": [xs[0], y, xs[0] + 70, y + 20]})
        for c, text in enumerate(values(r), start=1):
            if text:
                out.append({"text": text, "bbox": [xs[c], y, xs[c] + 50, y + 20]})
    return [*out, *(_footer(footer, 500) if footer else [])]


def _code_batch(tmp_path, first, second=None):
    template = tmp_path / "ma.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A3"], ws["B3"], ws["C3"] = "Mã", "Số lượng", "Đơn giá"
    for r, code in enumerate(CODES, start=4):
        ws[f"A{r}"] = code
    wb.save(template)
    source = _source(first, copy.deepcopy(first) if second is None else second)
    guess = infer_spec(read_source(template), source, frozenset(), experimental_image_sources=True)
    assert guess.spec is not None, [q.code for q in guess.questions]
    form = read_form(template, guess.spec)
    return propose(form, source, plan_headers(form, source), experimental_image_sources=True)


def test_table_of_text_keys_and_integers_is_found_on_an_image(tmp_path):
    batch = _code_batch(tmp_path, _code_tokens(lambda r: (str(10 + r), str(1000 + r))))
    assert _colors(batch) == {"yellow": 16}
    assert {c.target.a1_at_write: c.value.normalized for c in batch.changes}["B4"] == "10"


def test_a_data_row_is_never_taken_as_the_header_on_an_image(tmp_path):
    """Số ở dạng chữ từng làm dòng dữ liệu trông như tiêu đề. Có ô trống ở giữa bảng thì tiêu đề giả đó còn thành bảng."""
    batch = _code_batch(tmp_path, _code_tokens(lambda r: (str(10 + r), "" if r == 5 else str(1000 + r))))
    assert {c.mapping.source_header for c in batch.changes} == {"Số lượng", "Đơn giá"}


def test_printed_total_is_reached_on_a_table_of_text_keys(tmp_path):
    batch = _code_batch(tmp_path, _code_tokens(lambda r: (str(10 + r), str(1000 + r)), "Tổng số mẫu: 23"))
    assert _colors(batch) == {"red": 16}
    assert all("CHECK_PRINTED_TOTAL_MISMATCH" in _codes(c) for c in batch.changes)


def test_missing_cells_are_not_reported_as_a_wrong_total(tmp_path):
    batch = _code_batch(tmp_path, _code_tokens(lambda r: (str(10 + r), "" if r in (2, 3) else str(1000 + r)), "Tổng số mẫu: 16"))
    assert _colors(batch)["red"] == 14
    assert all("CHECK_TOTAL_NOT_COMPARABLE" in _codes(c) and "CHECK_PRINTED_TOTAL_MISMATCH" not in _codes(c) for c in batch.changes)
    notice = next(n for n in batch.notices if n.reason == "CHECK_TOTAL_NOT_COMPARABLE")
    assert "đọc thiếu" in notice.detail and "chưa chắc sai" in notice.detail


def test_full_table_with_another_total_is_reported_as_a_wrong_total(tmp_path):
    batch = _code_batch(tmp_path, _code_tokens(lambda r: (str(10 + r), str(1000 + r)), "Tổng số mẫu: 23"))
    notice = next(n for n in batch.notices if n.reason == "CHECK_PRINTED_TOTAL_MISMATCH")
    assert "app đếm được" in notice.detail and "xem lại phiếu gốc" in notice.detail


def test_column_total_with_missing_cells_is_not_comparable(tmp_path):
    xs = [50, 200, 400]
    tokens = _code_tokens(lambda r: (str(10 + r), "" if r == 2 else str(1000 + r)))
    y = 130 + 30 * len(CODES)
    tokens += [{"text": "Tổng", "bbox": [xs[0], y, xs[0] + 70, y + 20]}, {"text": "999", "bbox": [xs[1], y, xs[1] + 50, y + 20]}, {"text": "9999", "bbox": [xs[2], y, xs[2] + 50, y + 20]}]
    batch = _code_batch(tmp_path, tokens)
    by_col = defaultdict(set)
    for c in batch.changes:
        by_col[c.target.a1_at_write[0]] |= _codes(c)
    assert "CHECK_PRINTED_TOTAL_MISMATCH" in by_col["B"] and "CHECK_TOTAL_NOT_COMPARABLE" in by_col["C"]


def test_excel_numbers_stored_as_text_are_still_labels_as_before(tmp_path):
    """Sửa cho ảnh, PDF không được đổi đường Excel: ô Excel không có `read`."""
    from vmai.infer import _is_label

    assert _is_label(SourceCell(a1="A1", kind="text", value="123"))
    assert not _is_label(SourceCell(a1="A1", kind="text", value="123", read=CellRead(page=1, bbox=(0, 0, 1, 1), reader="ocr:a@1")))
    assert _is_label(SourceCell(a1="A1", kind="text", value="Mã", read=CellRead(page=1, bbox=(0, 0, 1, 1), reader="ocr:a@1")))


def test_integer_column_with_one_much_larger_value_is_not_a_decimal_shift(tmp_path):
    """Đơn giá `50000`, `52000` với một món `2000` là số thật. Từng báo nhầm trên bảng vật tư tự dựng."""
    batch = _code_batch(tmp_path, _code_tokens(lambda r: (str(10 + r), "2000" if r == 3 else str(50000 + r * 1000))))
    assert not any("CHECK_DECIMAL_SHIFT" in _codes(c) for c in batch.changes)


def test_thousands_dot_with_no_evidence_stays_ambiguous_on_an_image(tmp_path):
    """`74.000` có thể là 74 hoặc 74 nghìn. Không có bằng chứng thì đỏ, không đoán theo cỡ số."""
    batch = _code_batch(tmp_path, _code_tokens(lambda r: (str(10 + r), f"{50 + r}.000")))
    prices = [c for c in batch.changes if c.target.a1_at_write.startswith("C")]
    assert prices and all(c.confidence.final_color is Color.RED and "READ_NUMBER_AMBIGUOUS" in _codes(c) for c in prices)


# ---- Tiêu đề hai tầng trên ảnh. Ghép sai tầng là số vào nhầm cột mà vẫn trông hợp lý, nên phần lớn ở đây là bẫy.


GROUPS = ("Natri", "Kali", "Clo")
XS_TIER = [50, 250, 400, 550, 700, 850, 1000]
TIER_DAYS = [f"{d:02d}/09/2026" for d in range(1, 7)]


def _tier_template(path, groups=GROUPS):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.merge_cells("A3:A4")
    ws["A3"] = "Ngày"
    for g, name in enumerate(groups):
        ws.cell(3, 2 + g * 2, name)
        ws.merge_cells(start_row=3, start_column=2 + g * 2, end_row=3, end_column=3 + g * 2)
        ws.cell(4, 2 + g * 2, "Mức 1")
        ws.cell(4, 3 + g * 2, "Mức 2")
    for r in range(6):
        ws.cell(5 + r, 1, dt.date(2026, 9, 1 + r))
    wb.save(path)
    return path


def _tier_tokens(uppers=None, upper_x=None, lowers=None):
    uppers = uppers or {i: name for i, name in enumerate(GROUPS)}
    lowers = lowers or ["Mức 1", "Mức 2"] * 3
    out = []
    for g, name in uppers.items():
        x = upper_x[g] if upper_x else (XS_TIER[1 + g * 2] + XS_TIER[2 + g * 2] + 60) / 2 - 30
        out.append({"text": name, "bbox": [x, 70, x + 60, 90]})
    out.append({"text": "Ngày", "bbox": [XS_TIER[0], 100, XS_TIER[0] + 60, 120]})
    for c, label in enumerate(lowers, start=1):
        out.append({"text": label, "bbox": [XS_TIER[c], 100, XS_TIER[c] + 60, 120]})
    for r, day in enumerate(TIER_DAYS):
        y = 130 + r * 30
        out.append({"text": day, "bbox": [XS_TIER[0], y, XS_TIER[0] + 110, y + 20]})
        for c in range(1, 7):
            # Mỗi cột một cỡ số riêng, để nhầm cột hiện ra ngay ở giá trị.
            out.append({"text": f"{c}{c},{r}", "bbox": [XS_TIER[c], y, XS_TIER[c] + 50, y + 20]})
    return out


def _tier_run(tmp_path, first, second=None, groups=GROUPS):
    template = _tier_template(tmp_path / "tang.xlsx", groups)
    source = _source(first, copy.deepcopy(first) if second is None else second)
    guess = infer_spec(read_source(template), source, frozenset(), experimental_image_sources=True)
    if guess.spec is None:
        return guess, None
    form = read_form(template, guess.spec)
    return guess, propose(form, source, plan_headers(form, source), experimental_image_sources=True)


def _letters(batch):
    return {c.target.a1_at_write[0]: c.value.normalized[:2] for c in batch.changes}


def test_two_tier_header_maps_each_column_to_its_own_group(tmp_path):
    guess, batch = _tier_run(tmp_path, _tier_tokens())
    assert guess.status == "ok" and _colors(batch) == {"yellow": 36}
    assert _letters(batch) == {"B": "11", "C": "22", "D": "33", "E": "44", "F": "55", "G": "66"}


def test_trap_uppers_in_another_order_follow_the_page_not_the_form(tmp_path):
    guess, batch = _tier_run(tmp_path, _tier_tokens(uppers={0: "Kali", 1: "Natri", 2: "Clo"}))
    assert _letters(batch) == {"B": "33", "C": "44", "D": "11", "E": "22", "F": "55", "G": "66"}


def test_trap_second_engine_reads_another_upper_word_group_is_not_mapped(tmp_path):
    second = _tier_tokens(uppers={0: "Natri", 1: "Clo", 2: "Clo"})
    guess, batch = _tier_run(tmp_path, _tier_tokens(), second)
    # Nhóm đọc lệch còn tên trơn `Mức 1`, khớp với mọi nhóm: thà hỏi cả bảng hơn đoán.
    assert batch is None or {"D", "E"}.isdisjoint(_letters(batch))
    assert batch is not None or guess.questions


def test_trap_second_engine_misses_the_upper_tier_nothing_is_composed(tmp_path):
    second = [t for t in _tier_tokens() if t["text"] not in GROUPS]
    guess, batch = _tier_run(tmp_path, _tier_tokens(), second)
    assert batch is None and guess.questions


def test_trap_uneven_groups_are_not_guessed(tmp_path):
    uppers_x = {0: XS_TIER[1], 1: XS_TIER[2], 2: XS_TIER[5]}
    guess, batch = _tier_run(tmp_path, _tier_tokens(upper_x=uppers_x))
    assert batch is None and guess.questions


def test_trap_upper_word_midway_between_two_columns_is_not_guessed(tmp_path):
    middle = (XS_TIER[2] + XS_TIER[3] + 60) / 2 - 30
    guess, batch = _tier_run(tmp_path, _tier_tokens(uppers={0: "Natri", 1: "Kali"}, upper_x={0: XS_TIER[1], 1: middle}))
    assert batch is None and guess.questions


def test_trap_lower_labels_that_differ_between_groups_are_not_composed(tmp_path):
    lowers = ["Mức 1", "Mức 2", "Mức 2", "Mức 1", "Mức 1", "Mức 2"]
    guess, batch = _tier_run(tmp_path, _tier_tokens(lowers=lowers))
    assert batch is None and guess.questions


def test_trap_group_name_not_on_the_form_is_not_filled(tmp_path):
    guess, batch = _tier_run(tmp_path, _tier_tokens(uppers={0: "Natri", 1: "Canxi", 2: "Clo"}))
    assert batch is not None and {"D", "E"}.isdisjoint(_letters(batch))


def test_trap_page_title_above_is_not_taken_as_an_upper_tier(tmp_path):
    tokens = [t for t in _tier_tokens() if t["text"] not in GROUPS]
    tokens.append({"text": "[MẪU GIẢ] BIỂU MẪU", "bbox": [300, 70, 700, 90]})
    guess, batch = _tier_run(tmp_path, tokens)
    assert batch is None and guess.questions


def test_upper_tier_with_another_unit_is_red(tmp_path):
    template = tmp_path / "don_vi.xlsx"
    _tier_template(template, ("Natri (mmol/L)", "Kali", "Clo"))
    tokens = _tier_tokens(uppers={0: "Natri (mg/dL)", 1: "Kali", 2: "Clo"})
    source = _source(tokens, copy.deepcopy(tokens))
    guess = infer_spec(read_source(template), source, frozenset(), experimental_image_sources=True)
    form = read_form(template, guess.spec)
    batch = propose(form, source, plan_headers(form, source), experimental_image_sources=True)
    natri = [c for c in batch.changes if c.target.a1_at_write[0] in "BC"]
    assert natri and all(c.confidence.final_color is Color.RED and "READ_UNIT_MISMATCH" in _codes(c) for c in natri)


def test_header_without_repeated_names_is_left_alone(tmp_path):
    from vmai.table_read import compose_two_tier, tokens_to_grid

    tokens = _tokens()
    grid = tokens_to_grid(tokens)[0]
    assert compose_two_tier(grid, tokens, tokens) is grid


def test_fuzzy_match_never_crosses_a_level_number():
    """Trên bảng hai tầng, cột `Mức 1` đọc lệch thì không được chiếm chỗ `Mức 2` đã ghép."""
    fields = {"a": ("Cholesterol Mức 2",)}
    assert match_field("Cholesterol Mức 1", fields, {}).field is None
    assert match_field("Cholesterol Múc 2", fields, {}).field == "a"


def test_trap_level_one_never_fills_a_level_two_column(tmp_path):
    """Một engine đọc lệch tên `Mức 2` nên chỉ cột `Mức 1` ở lại: số của nó không được vào cột khác."""
    first = _tier_tokens()
    second = _tier_tokens(lowers=["Mức 1", "Mức 9", "Mức 1", "Mức 2", "Mức 1", "Mức 2"])
    guess, batch = _tier_run(tmp_path, first, second)
    letters = _letters(batch)
    assert letters["B"] == "11" and "C" not in letters
    assert all(c.confidence.final_color is Color.YELLOW for c in batch.changes)


@pytest.mark.parametrize("word", ["Kali", "Hill", "Nil"])
def test_repair_header_leaves_ordinary_words_alone(word):
    from vmai.infer import repair_header

    assert repair_header(word) == word.lower()


def test_two_numbers_close_together_stay_in_two_cells():
    from vmai.table_read import _merge_cells

    cells = _merge_cells([{"text": "1,5", "bbox": [0, 0, 40, 20]}, {"text": "2,25", "bbox": [50, 0, 100, 20]}], 18)
    assert [c["text"] for c in cells] == ["1,5", "2,25"]


def test_number_with_a_space_for_thousands_stays_one_number():
    from vmai.table_read import _merge_cells

    cells = _merge_cells([{"text": "1", "bbox": [0, 0, 10, 20]}, {"text": "234", "bbox": [14, 0, 44, 20]}], 18)
    assert [c["text"] for c in cells] == ["1 234"]


# ---- Ô công thức có đầu vào trống, bị bỏ qua, còn chờ. Đề xuất, chờ PM duyệt.


def _formula(colors, **kw):
    from vmai import grade_formula

    return grade_formula(colors, False, **kw)


def test_empty_input_means_not_green_and_says_how_many():
    color, reasons = _formula([Color.GREEN] * 28, empty_inputs=2, total_inputs=30)
    assert color is Color.YELLOW
    assert [r.code for r in reasons] == ["FORMULA_INPUT_EMPTY"] and "28/30" in reasons[0].detail


def test_skipped_input_means_not_green_even_with_only_typed_inputs_left():
    color, reasons = _formula([None] * 29, skipped_inputs=1, total_inputs=30)
    assert color is Color.YELLOW and "FORMULA_INPUT_SKIPPED" in {r.code for r in reasons}


def test_a_pending_red_proposal_stays_red_and_is_not_raised_to_yellow():
    color, reasons = _formula([Color.GREEN, Color.RED], pending_inputs=1, total_inputs=2)
    assert color is Color.RED and "FORMULA_INPUT_PENDING" in {r.code for r in reasons}


def test_a_pending_green_proposal_does_not_leave_the_formula_green():
    """Số chưa vào lưới nên công thức đang tính thiếu ô đó, dù đề xuất xanh."""
    color, reasons = _formula([Color.GREEN] * 30, pending_inputs=1, total_inputs=30)
    assert color is Color.YELLOW and "29/30" in reasons[0].detail


def test_only_typed_inputs_and_empty_cells_still_have_no_color():
    color, reasons = _formula([None, None], empty_inputs=5, total_inputs=7)
    assert color is None and not reasons


def test_no_missing_input_changes_nothing():
    assert _formula([Color.GREEN, Color.GREEN], total_inputs=2) == (Color.GREEN, ())


@pytest.mark.skipif(not FULL_TREE, reason="Kho công khai không có docs/ nội bộ: dùng bản chụp tests/data/doc_codes.json")
def test_doc_snapshot_matches_the_docs():
    """Trên cây đầy đủ không bỏ qua: thiếu docs/ thì lỗi, bản chụp cũ thì lỗi."""
    assert DOC.exists(), "Cây đầy đủ phải có docs/ai/confidence_rules.md"
    fresh = doc_snapshot.build()
    assert fresh == doc_snapshot.load(), "Bản chụp cũ: chạy vmai/tests/doc_snapshot.py"
