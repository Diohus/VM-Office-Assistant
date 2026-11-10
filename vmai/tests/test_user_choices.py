"""Cột do người dùng chọn (US-1.13, T-AI-06), giá trị ví dụ, lô xuất làm fixture. Dữ liệu giả, tự dựng."""
from __future__ import annotations

import datetime as dt
import json
import re
import sys
from pathlib import Path

import doc_snapshot
import openpyxl
import pytest

import vmai
from vmai import Color, confirmation_rows, infer_spec, plan_headers, propose, read_form, read_source, source_options
from vmai.contract import ProposalBatch
from vmai.infer import SpecGuess, profile_from_guess
from vmai.profile import TemplateProfile
from vmai.template import sha256_file

ROOT = Path(__file__).resolve().parents[2]
X01 = ROOT / "fixtures" / "eval" / "generated" / "xlsx" / "X01_m0"
DAYS = [dt.datetime(2026, 9, d) for d in range(1, 7)]
WHEN = "2026-11-05T00:00:00+07:00"


def _book(path, rows):
    wb = openpyxl.Workbook()
    for row in rows:
        wb.active.append(row)
    wb.save(path)
    return path


def _template(tmp_path, headers=("Ngày", "Natri")):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "[MẪU GIẢ]"
    for i, h in enumerate(headers, start=1):
        ws.cell(3, i, h)
    for r, day in enumerate(DAYS, start=4):
        ws.cell(r, 1, day)
    path = tmp_path / "t.xlsx"
    wb.save(path)
    return path


def _source(tmp_path, name="s.xlsx"):
    return _book(tmp_path / name, [["Ngày", "Cột A", "Cột B", "Người thực hiện"], *[[d, 140 + i, 4.1 + i / 10, "Nhân viên giả"] for i, d in enumerate(DAYS)]])


def _codes(change):
    return {r.code for r in change.confidence.reasons}


def _run(template, source, guess, profile=None):
    form = read_form(template, guess.spec)
    return form, propose(form, source, plan_headers(form, source, profile))


def test_unmapped_column_is_offered_with_its_letter(tmp_path):
    guess = infer_spec(read_source(_template(tmp_path)), read_source(_source(tmp_path)))
    assert guess.spec is None and "COLUMN_NOT_MAPPED" in {q.code for q in guess.questions}
    assert [(c.letter, c.col_path) for c in guess.open_columns] == [("B", ("Natri",))]


def test_user_choice_fills_the_chosen_column_at_most_yellow(tmp_path):
    template, source = _template(tmp_path), read_source(_source(tmp_path))
    guess = infer_spec(read_source(template), source, column_choices={"B": "Cột A"})
    assert guess.status == "ok" and guess.columns[0].chosen_by_user and guess.columns[0].source_header == "Cột A"
    _, batch = _run(template, source, guess)
    assert [c.value.normalized for c in batch.changes] == [str(140 + i) for i in range(6)]
    assert all(c.confidence.final_color is Color.YELLOW and "MAP_COLUMN_USER_CHOSEN" in _codes(c) for c in batch.changes)


def test_choice_wins_over_the_automatic_match_but_not_silently(tmp_path):
    template = _template(tmp_path, ("Ngày", "Natri"))
    source_path = _book(tmp_path / "s2.xlsx", [["Ngày", "Natri", "Khác"], *[[d, 140, 999] for d in DAYS]])
    source = read_source(source_path)
    assert infer_spec(read_source(template), source).columns[0].source_header == "Natri"
    guess = infer_spec(read_source(template), source, column_choices={"B": "Khác"})
    _, batch = _run(template, source, guess)
    assert {c.value.normalized for c in batch.changes} == {"999"}
    assert all("MAP_COLUMN_USER_CHOSEN" in _codes(c) for c in batch.changes)


def test_choosing_blank_leaves_the_column_out(tmp_path):
    template = _template(tmp_path, ("Ngày", "Natri", "Kali"))
    source = read_source(_book(tmp_path / "s3.xlsx", [["Ngày", "Natri", "Kali"], *[[d, 140, 4] for d in DAYS]]))
    guess = infer_spec(read_source(template), source, column_choices={"C": None})
    _, batch = _run(template, source, guess)
    assert {c.target.a1_at_write[0] for c in batch.changes} == {"B"}


@pytest.mark.parametrize("bad", ["Người thực hiện", "Ngày", "Không có cột này"], ids=["cot_ve_nguoi", "cot_khoa", "khong_co"])
def test_a_choice_that_cannot_be_used_is_refused(tmp_path, bad):
    template, source = _template(tmp_path), read_source(_source(tmp_path))
    guess = infer_spec(read_source(template), source, column_choices={"B": bad})
    assert guess.spec is None and "COLUMN_CHOICE_INVALID" in {q.code for q in guess.questions}


def test_a_choice_for_a_column_the_form_does_not_fill_is_refused(tmp_path):
    guess = infer_spec(read_source(_template(tmp_path)), read_source(_source(tmp_path)), column_choices={"B": "Cột A", "Z": "Cột B"})
    assert guess.spec is None and "COLUMN_CHOICE_INVALID" in {q.code for q in guess.questions}


def test_same_source_column_for_two_form_columns_is_refused(tmp_path):
    template = _template(tmp_path, ("Ngày", "Natri", "Kali"))
    guess = infer_spec(read_source(template), read_source(_source(tmp_path)), column_choices={"B": "Cột A", "C": "Cột A"})
    assert guess.spec is None and "COLUMN_MAPPING_CONFLICT" in {q.code for q in guess.questions}


def test_two_source_columns_with_the_same_name_cannot_be_chosen(tmp_path):
    source = read_source(_book(tmp_path / "s4.xlsx", [["Ngày", "Giá trị", "Giá trị"], *[[d, 1, 2] for d in DAYS]]))
    guess = infer_spec(read_source(_template(tmp_path)), source, column_choices={"B": "Giá trị"})
    assert guess.spec is None and "COLUMN_CHOICE_INVALID" in {q.code for q in guess.questions}
    assert all(not option.usable for option in source_options(source))


def test_user_choice_with_another_unit_is_still_red(tmp_path):
    template = _template(tmp_path, ("Ngày", "HGB (g/L)"))
    source = read_source(_book(tmp_path / "s5.xlsx", [["Ngày", "Hemoglobin (g/dL)"], *[[d, 13.5] for d in DAYS]]))
    guess = infer_spec(read_source(template), source, column_choices={"B": "Hemoglobin (g/dL)"})
    _, batch = _run(template, source, guess)
    assert all(c.confidence.final_color is Color.RED and "READ_UNIT_MISMATCH" in _codes(c) for c in batch.changes)


# ---- Cột do người chọn rồi lưu hồ sơ: đường duy nhất có thể lên xanh (Q17).


def _profile(guess, form, plan_source):
    plan = plan_headers(form, plan_source)
    return profile_from_guess(guess, form.fingerprint, plan.headers, confirmed_by="tester.gia", confirmed_at=WHEN, profile_id="gia")


def test_chosen_column_without_limits_stays_yellow_next_month(tmp_path):
    template, source = _template(tmp_path), read_source(_source(tmp_path))
    guess = infer_spec(read_source(template), source, column_choices={"B": "Cột A"})
    form = read_form(template, guess.spec)
    profile = _profile(guess, form, source)
    batch = propose(form, source, plan_headers(form, source, profile))
    assert batch.changes and all(c.confidence.final_color is Color.YELLOW for c in batch.changes)
    assert all("MAP_PROFILE_USER_ANSWER" in _codes(c) for c in batch.changes)  # Q115: cột tự chọn, trần vàng có trước Q17


def _qc_sources(tmp_path):
    """Nguồn dạng dọc của X01, đổi tên cột kết quả, thêm một cột số sai gấp 10."""
    wb = openpyxl.load_workbook(X01 / "source.xlsx")
    ws = wb.worksheets[0]
    header_row = next(r for r in range(1, 10) if any(str(ws.cell(r, c).value or "").strip() in ("Kết quả", "KQ", "Ket qua") for c in range(1, ws.max_column + 1)))
    result = next(c for c in range(1, ws.max_column + 1) if str(ws.cell(header_row, c).value or "").strip() in ("Kết quả", "KQ", "Ket qua"))
    ws.cell(header_row, result, "Giá trị đo")
    extra = ws.max_column + 1
    ws.cell(header_row, extra, "Giá trị khác")
    for r in range(header_row + 1, ws.max_row + 1):
        value = ws.cell(r, result).value
        if isinstance(value, (int, float)):
            ws.cell(r, extra, value * 10)
    path = tmp_path / "qc.xlsx"
    wb.save(path)
    return read_source(path)


def _qc_choices(guess):
    return guess.open_columns


def test_right_column_chosen_on_a_form_with_limits_stays_yellow_even_after_the_profile(tmp_path):
    """Q115: trước đây lần sau lên xanh. Chọn nhầm cột gần giống (máy 2) thì giới hạn không bắt được, nên cột tự chọn tối đa vàng mãi."""
    template = X01 / "template.xlsx"
    source = _qc_sources(tmp_path)
    first = infer_spec(read_source(template), source, frozenset({"LEVEL_NAME_DIFFERS"}))
    assert first.spec is None and first.open_columns, [q.code for q in first.questions]
    choices = {c.letter: "Giá trị đo" for c in first.open_columns}
    guess = infer_spec(read_source(template), source, frozenset({"LEVEL_NAME_DIFFERS"}), column_choices=choices)
    assert guess.spec is not None, [q.detail for q in guess.questions]
    form = read_form(template, guess.spec)
    now = propose(form, source, plan_headers(form, source))
    assert now.changes and Color.GREEN not in {c.confidence.final_color for c in now.changes}  # lô đang duyệt không xanh
    later = propose(form, source, plan_headers(form, source, _profile(guess, form, source)))
    assert later.changes and Color.GREEN not in {c.confidence.final_color for c in later.changes}
    assert all("MAP_PROFILE_USER_ANSWER" in _codes(c) for c in later.changes)


def test_trap_wrong_column_chosen_and_saved_never_goes_green(tmp_path):
    """Đường số sai còn lại: người chọn nhầm cột rồi lưu hồ sơ. Cột có giới hạn phải tự bắt được, không xanh ô nào."""
    template = X01 / "template.xlsx"
    source = _qc_sources(tmp_path)
    first = infer_spec(read_source(template), source, frozenset({"LEVEL_NAME_DIFFERS"}))
    choices = {c.letter: "Giá trị khác" for c in first.open_columns}
    guess = infer_spec(read_source(template), source, frozenset({"LEVEL_NAME_DIFFERS"}), column_choices=choices)
    form = read_form(template, guess.spec)
    later = propose(form, source, plan_headers(form, source, _profile(guess, form, source)))
    limited = [c for c in later.changes if form.columns and next(k for k in form.columns if k.letter == c.target.a1_at_write[0]).limit is not None]
    assert limited, "ca này phải có cột có giới hạn"
    assert Color.GREEN not in {c.confidence.final_color for c in limited}
    assert Color.GREEN not in {c.confidence.final_color for c in later.changes}


def test_known_gap_wrong_column_chosen_without_limits_stays_yellow_and_wrong(tmp_path):
    """Rủi ro còn lại, cùng loại với L24: không có giới hạn thì app không biết người chọn nhầm. Vàng, người duyệt phải thấy."""
    template, source = _template(tmp_path), read_source(_source(tmp_path))
    guess = infer_spec(read_source(template), source, column_choices={"B": "Cột B"})
    form = read_form(template, guess.spec)
    batch = propose(form, source, plan_headers(form, source, _profile(guess, form, source)))
    assert {c.confidence.final_color for c in batch.changes} == {Color.YELLOW}


# ---- Giá trị ví dụ


def test_confirmation_rows_have_samples_and_open_columns(tmp_path):
    template = _template(tmp_path, ("Ngày", "Natri", "Kali"))
    source = read_source(_book(tmp_path / "s6.xlsx", [["Ngày", "Natri", "Cột lạ"], *[[d, 140 + i, 1] for i, d in enumerate(DAYS)]]))
    guess = infer_spec(read_source(template), source)
    rows = {r.letter: r for r in confirmation_rows(guess, source)}
    assert rows["B"].source_header == "Natri" and rows["B"].samples == ("140", "141", "142")
    assert rows["C"].source_header is None and rows["C"].samples == ()


def test_source_options_leave_out_people_and_dates(tmp_path):
    options = source_options(read_source(_source(tmp_path)))
    assert [o.header for o in options] == ["Cột A", "Cột B"]
    assert options[0].samples == ("140", "141", "142")


def test_samples_of_a_long_source_belong_to_that_test_and_level(tmp_path):
    template = X01 / "template.xlsx"
    source = read_source(X01 / "source.xlsx")
    guess = infer_spec(read_source(template), source, frozenset({"LEVEL_NAME_DIFFERS"}))
    rows = confirmation_rows(guess, source)
    with_dims = [r for r in rows if r.dims]
    assert with_dims and all(r.samples for r in with_dims)
    assert len({r.samples for r in with_dims}) > 1, "các cột khác nhau phải có ví dụ khác nhau"


def test_person_like_value_is_not_a_sample(tmp_path):
    source = read_source(_book(tmp_path / "s7.xlsx", [["Ngày", "Natri"], [DAYS[0], "Nguyễn Văn A"], *[[d, 140] for d in DAYS[1:]]]))
    assert all("Nguyễn" not in s for o in source_options(source) for s in o.samples)


def test_the_samples_module_does_not_touch_the_llm_layer():
    text = (Path(vmai.__file__).parent / "confirm.py").read_text(encoding="utf-8")
    assert "llm" not in re.sub(r'""".*?"""', "", text, flags=re.S).lower()


# ---- Lô xuất làm fixture (T-AI-06, T-AI-07)


@pytest.fixture(scope="module")
def exported(tmp_path_factory):
    sys.path.insert(0, str(ROOT / "vmai" / "eval"))
    import export_batches

    target = tmp_path_factory.mktemp("lo")
    return target, export_batches.build(target), export_batches


def _load(target, entry):
    return ProposalBatch.model_validate_json((target / entry["file"]).read_text(encoding="utf-8"))


def _by_file(index):
    return {e["file"]: e for e in index["batches"]}


def test_export_has_every_color_and_every_batch_is_valid(exported):
    target, index, _ = exported
    colors = set()
    for entry in index["batches"]:
        batch = _load(target, entry)
        assert batch.rule_version == vmai.RULE_VERSION and batch.proposal_id.startswith("mau-")
        colors |= set(entry["colors"])
    assert colors == {"red", "yellow", "green"}


def test_every_batch_comes_with_files_whose_sha_matches(exported):
    """Dev: thiếu cái này thì kiểm hash và mở nguồn (US-1.8) đều hỏng."""
    target, index, _ = exported
    for entry in index["batches"]:
        batch = _load(target, entry)
        assert sha256_file(target / entry["template"]) == batch.template.template_sha256, entry["file"]
        assert sha256_file(target / entry["source"]) == batch.sources[0].source_file_sha256, entry["file"]


@pytest.mark.parametrize(
    ("name", "wanted"),
    [("X01_m0", "green"), ("X01b", "yellow"), ("X04", "red"), ("X05", "red")],
)
def test_cases_dev_asked_for_are_there(exported, name, wanted):
    _, index, _ = exported
    assert wanted in _by_file(index)[f"{name}.json"]["colors"]


def test_x11_has_cells_with_existing_numbers_reviewed_one_by_one(exported):
    target, index, _ = exported
    batch = _load(target, _by_file(index)["X11.json"])
    existing = [c for c in batch.changes if c.previous.value is not None]
    assert existing and all(c.review.unit == "cell" for c in existing)


def test_a_batch_has_a_red_cell_inside_the_statistic_range_and_the_form_has_the_formula(exported):
    target, index, _ = exported
    entry = _by_file(index)[index["red_in_stat_range"]]
    batch = _load(target, entry)
    assert any(c.confidence.final_color is Color.RED and re.fullmatch(r"B(?:[89]|[12]\d|3[0-7])", c.target.a1_at_write) for c in batch.changes)
    sheet = openpyxl.load_workbook(target / entry["template"])[batch.changes[0].target.sheet]
    assert "B8:B37" in str(sheet["B39"].value)


def test_image_batch_has_a_page_image_that_matches_for_cropping(exported):
    target, index, _ = exported
    entry = _by_file(index).get("anh_co_trang.json")
    assert entry is not None, "Thiếu lô ảnh: token chép sẵn ở fixtures/eval/ocr_tokens/clean_s301.json"
    source = _load(target, entry).sources[0]
    assert len(source.pages) == 1 and source.pages[0].page_image_sha256 == sha256_file(target / entry["page_image"])
    assert "yellow" in entry["colors"]


def test_us113_data_loads_back_and_the_profile_turns_the_batch_green(exported):
    target, index, _ = exported
    data = index["us113"]
    guess = SpecGuess.model_validate_json((target / data["doan_spec_guess"]).read_text(encoding="utf-8"))
    profile = TemplateProfile.model_validate_json((target / data["doan_ho_so"]).read_text(encoding="utf-8"))
    assert guess.spec is not None and guess.status == "ok" and profile.confirmed_by
    assert set(_by_file(index)["us113_lan_dau.json"]["colors"]) == {"yellow"}
    assert "green" in _by_file(index)["us113_lan_sau.json"]["colors"]
    screen = json.loads((target / data["doan_man_xac_nhan"]).read_text(encoding="utf-8"))
    assert screen["rows"] and all(r["samples"] for r in screen["rows"] if r["source_header"])


def test_user_chosen_column_has_a_yellow_batch_and_wrong_choice_is_caught(exported):
    target, index, _ = exported
    by_file = _by_file(index)
    assert by_file["nguoi_chon_cot.json"]["colors"].get("yellow") and "green" not in by_file["nguoi_chon_cot.json"]["colors"]
    assert "MAP_COLUMN_USER_CHOSEN" in by_file["nguoi_chon_cot.json"]["reasons"]
    later = by_file["nguoi_chon_cot_lan_sau.json"]
    assert later["colors"].get("yellow") and "green" not in later["colors"]  # Q115
    assert "MAP_PROFILE_USER_ANSWER" in later["reasons"]
    wrong = by_file["nguoi_chon_nham_cot.json"]["colors"]
    assert wrong.get("red") and "green" not in wrong
    asked = SpecGuess.model_validate_json((target / index["us113"]["chon_cot_hoi_lai"]).read_text(encoding="utf-8"))
    assert asked.spec is None and asked.open_columns


def test_stale_pair_changes_only_what_it_says(exported):
    target, index, _ = exported
    stale = index["stale"]
    filled = openpyxl.load_workbook(target / stale["bieu_mau_da_dien"])
    edited = openpyxl.load_workbook(target / stale["X14"]["file"])
    sheet = stale["X14"]["sheet"]
    differ = [c.coordinate for row in filled[sheet].iter_rows() for c in row if c.value != edited[sheet][c.coordinate].value]
    assert sorted(differ) == sorted(stale["X14"]["cells"]) and len(differ) == 3
    assert stale["X15"]["sha256"] != stale["X15"]["original_sha256"] == stale["lo_goc"]["source_file_sha256"]
    assert stale["X15"]["expected_reason"] == "STALE_SOURCE_CHANGED" and stale["X14"]["expected_reason"] == "STALE_VALUE_CHANGED"
    assert stale["X14b"]["cells"], "biểu mẫu phải có ô công thức để sửa"


def test_export_has_questions_and_errors(exported):
    _, index, _ = exported
    assert {"COLUMN_NOT_MAPPED", "KEY_NOT_FOUND"} <= {q["code"] for g in index["guesses"] for q in g["questions"]}
    assert any(g["status"] == "ok" and g["confirmation_rows"] for g in index["guesses"])
    assert {"PDF_TEXT_LAYER_UNTRUSTED", "FEATURE_NOT_RELEASED", "FILE_PASSWORD"} <= {e["code"] for e in index["errors"]}


def test_export_is_stable_and_has_no_person_values(exported, tmp_path):
    target, index, module = exported
    again = module.build(tmp_path / "lai")
    assert again == index, "xuất lại phải ra đúng file cũ, kể cả mã băm"
    for entry in index["batches"]:
        assert (target / entry["file"]).read_bytes() == (tmp_path / "lai" / entry["file"]).read_bytes()
    for path in target.glob("*.json"):
        body = path.read_text(encoding="utf-8")
        assert "Nguyễn Giả Một" not in body and "sk-" not in body


def test_export_never_lists_green_from_an_image_or_pdf(exported):
    _, index, _ = exported
    assert not [e for e in index["batches"] if e["kind"] != "xlsx" and "green" in e["colors"]]


# ---- Bảng câu tiếng thường


def test_every_code_in_the_source_has_a_plain_sentence():
    table = doc_snapshot.load()["reason_texts"]  # bản chụp từ docs/ai/reason_texts.md, test_rules kiểm đồng bộ
    source = "\n".join(p.read_text(encoding="utf-8") for p in Path(vmai.__file__).parent.glob("*.py"))
    codes = set(re.findall(r'"([A-Z][A-Z0-9]*(?:_[A-Z0-9]+)+)"', source))
    not_user_facing = {"OMP_THREAD_LIMIT", "RULE_VERSION", "SCHEMA_VERSION", "VMAI_TESSDATA", "VMAI_TESSERACT", "VMAI_TESSERACT_DOCKER"}
    prefixes = {"READ_FORMULA", "READ_NOT", "READ_NUMBER", "READ_TEXT", "READ_UNIT", "SOURCE_DUPLICATE"}  # tiền tố để gộp thông báo
    missing = sorted(c for c in codes - not_user_facing - prefixes if c not in table)
    assert not missing, missing
    no_sentence = sorted(c for c in codes - not_user_facing - prefixes if not table[c])
    assert not no_sentence, f"Có mã nhưng chưa có câu hiện: {no_sentence}"


def _two_machines(tmp_path):
    """Nguồn X01 có hai cột "KQ máy 1", "KQ máy 2". Máy 2 lệch khoảng 1%, vẫn trong ±3SD. [MẪU GIẢ]"""
    import random

    wb = openpyxl.load_workbook(X01 / "source.xlsx")
    ws = wb.worksheets[0]
    headers = [c.value for c in ws[1]]
    result = headers.index("Kết quả") + 1
    ws.cell(1, result, "KQ máy 1")
    extra = ws.max_column + 1
    ws.cell(1, extra, "KQ máy 2")
    rng = random.Random(5)
    for r in range(2, ws.max_row + 1):
        value = ws.cell(r, result).value
        if isinstance(value, (int, float)):
            other = round(value * (1 + rng.uniform(-0.01, 0.01)), 2)
            ws.cell(r, extra, other if other != value else round(value + 0.01, 2))
    path = tmp_path / "hai_may.xlsx"
    wb.save(path)
    return read_source(path)


def test_trap_wrong_machine_chosen_and_saved_never_goes_green(tmp_path):
    """Q115. Trước khi sửa: lần sau 180/180 ô xanh, đều sai. Số máy 2 nằm trong giới hạn nên Q17 và CHECK_COLUMN_SUSPECT không bắt được."""
    template, source = X01 / "template.xlsx", _two_machines(tmp_path)
    first = infer_spec(read_source(template), source, frozenset({"LEVEL_NAME_DIFFERS"}))
    letters = [c.letter for c in first.open_columns]
    assert letters, [q.code for q in first.questions]
    right = infer_spec(read_source(template), source, frozenset({"LEVEL_NAME_DIFFERS"}), column_choices={k: "KQ máy 1" for k in letters})
    form_right = read_form(template, right.spec)
    truth = {c.target.a1_at_write: c.value.normalized for c in propose(form_right, source, plan_headers(form_right, source)).changes}
    wrong = infer_spec(read_source(template), source, frozenset({"LEVEL_NAME_DIFFERS"}), column_choices={k: "KQ máy 2" for k in letters})
    form = read_form(template, wrong.spec)
    later = propose(form, source, plan_headers(form, source, _profile(wrong, form, source)))
    assert len(later.changes) == 180 and sum(c.value.normalized != truth[c.target.a1_at_write] for c in later.changes) > 150
    assert Color.GREEN not in {c.confidence.final_color for c in later.changes}


def _l01_two_columns(tmp_path, right_header, wrong_header):
    """Nguồn L01, cột đúng đổi tên, thêm cột máy 2 lệch khoảng 1%. [MẪU GIẢ]"""
    import random

    layout = ROOT / "fixtures" / "eval" / "generated" / "layouts" / "L01_hai_dong_tieu_de"
    wb = openpyxl.load_workbook(layout / "source.xlsx")
    ws = wb.worksheets[0]
    result = [c.value for c in ws[1]].index("Kết quả") + 1
    ws.cell(1, result, right_header)
    extra = ws.max_column + 1
    ws.cell(1, extra, wrong_header)
    rng = random.Random(9)
    for r in range(2, ws.max_row + 1):
        value = ws.cell(r, result).value
        if isinstance(value, (int, float)):
            other = round(value * (1 + rng.uniform(-0.01, 0.01)), 2)
            ws.cell(r, extra, other if other != value else round(value + 0.01, 2))
    path = tmp_path / "l01_hai_cot.xlsx"
    wb.save(path)
    return layout, read_source(path), json.loads((layout / "answer.json").read_text(encoding="utf-8"))["cells"]


@pytest.mark.parametrize("right, wrong", [("KQ máy 1", "Kết quả"), ("Kết quả", "Kết quả máy 2")], ids=["doan_nham_may_2", "doan_dung"])
def test_guess_among_similar_columns_never_goes_green_after_saving(tmp_path, right, wrong):
    """T-AI-12. Bộ đoán tự chọn cột tên "Kết quả" mà không hỏi. Ca doan_nham_may_2 trước khi sửa: lần sau 40/40 ô xanh, đều sai."""
    layout, source, gold = _l01_two_columns(tmp_path, right, wrong)
    guess = infer_spec(read_source(layout / "template.xlsx"), source)
    assert guess.spec is not None and not guess.questions  # đường không hỏi
    form = read_form(layout / "template.xlsx", guess.spec)
    later = propose(form, source, plan_headers(form, source, _profile(guess, form, source)))
    assert len(later.changes) == 40 and Color.GREEN not in {c.confidence.final_color for c in later.changes}
    assert all("MAP_PROFILE_SIMILAR_COLUMNS" in _codes(c) for c in later.changes)
    if right == "KQ máy 1":
        assert all(c.value.normalized != gold[c.target.a1_at_write] for c in later.changes)  # đúng là đoán nhầm
