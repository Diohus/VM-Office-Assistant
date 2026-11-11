"""Hàm vào lát 1 `fill_from_excel` (Q133, T-AI-19). [MẪU GIẢ] Mọi tên người, mã người bệnh đều giả."""
from __future__ import annotations

import importlib.util
import json
import shutil
import uuid
from collections import Counter
from pathlib import Path

import openpyxl
import pytest

import vmai.fill as fill_module
from vmai import Color, FillError, fill_from_excel, load_spec, plan_headers, propose, read_form, read_source
from vmai.fill import _cap_green, user_message

ROOT = Path(__file__).resolve().parents[2]
X01 = ROOT / "fixtures" / "eval" / "generated" / "xlsx" / "X01_m0"
X01_M1 = ROOT / "fixtures" / "eval" / "generated" / "xlsx" / "X01_m1"
TEMPLATES = ROOT / "fixtures" / "templates"
PERSON_NAMES = ("Lê Giả A", "Phạm Giả B", "Hoàng Giả C")  # cột Người chạy của X01, sinh bởi generate.py


def _temp_copy(tmp_path: Path, path: Path) -> Path:
    """Như reader: chép ra thư mục tạm với tên uuid."""
    target = tmp_path / f"{uuid.uuid4().hex}{path.suffix}"
    shutil.copyfile(path, target)
    return target


def _fill(tmp_path, template=X01 / "template.xlsx", source=X01 / "source.xlsx", **kw):
    return fill_from_excel(_temp_copy(tmp_path, template), _temp_copy(tmp_path, source), spec_id="GIA-QC-01", **kw)


def _green_batch():
    """Lô có ô xanh thật: X01_m1 kèm hồ sơ đã xác nhận, như run_xlsx."""
    spec = importlib.util.spec_from_file_location("run_xlsx", ROOT / "vmai" / "eval" / "run_xlsx.py")
    run_xlsx = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(run_xlsx)
    answer = json.loads((X01_M1 / "answer.json").read_text(encoding="utf-8"))
    form = read_form(X01_M1 / "template.xlsx", run_xlsx.SPEC)
    source = read_source(X01_M1 / "source.xlsx")
    batch = propose(form, source, plan_headers(form, source, run_xlsx.load_profile(answer, form.fingerprint)))
    assert any(c.confidence.final_color is Color.GREEN for c in batch.changes), "ca dựng phải có ô xanh, nếu không test trần vàng vô nghĩa"
    return batch


# ---- Lô bình thường


def test_qc_source_gives_a_valid_batch_with_no_green(tmp_path):
    batch = _fill(tmp_path)
    assert len(batch.changes) == 180
    assert Counter(c.confidence.final_color for c in batch.changes) == {Color.YELLOW: 180}
    assert batch.template.spec_id == "GIA-QC-01"


def test_display_names_replace_temp_names_and_go_nowhere_else(tmp_path):
    template, source = _temp_copy(tmp_path, X01 / "template.xlsx"), _temp_copy(tmp_path, X01 / "source.xlsx")
    batch = fill_from_excel(template, source, spec_id="GIA-QC-01", template_name="Mau QC.xlsx", source_name="KQ_Nguyen_Gia_Thu.xlsx")
    assert batch.template.file_name == "Mau QC.xlsx" and [s.file_name for s in batch.sources] == ["KQ_Nguyen_Gia_Thu.xlsx"]
    dumped = batch.model_dump_json()
    assert template.stem not in dumped and source.stem not in dumped
    assert dumped.count("KQ_Nguyen_Gia_Thu") == 1  # chỉ ở sources[].file_name
    assert not [r for c in batch.changes for r in c.confidence.reasons if "Nguyen" in r.detail]


# ---- Trần vàng (Q133)


def test_default_cap_lowers_every_green_and_never_raises(monkeypatch, tmp_path):
    green = _green_batch()
    monkeypatch.setattr(fill_module, "propose", lambda *a, **k: green)
    capped = _fill(tmp_path)
    assert Color.GREEN not in {c.confidence.final_color for c in capped.changes}
    before = {c.change_id: c.confidence.final_color for c in green.changes}
    for change in capped.changes:
        was = before[change.change_id]
        if was is Color.GREEN:
            assert change.confidence.final_color is Color.YELLOW
            assert "MAP_GREEN_DISABLED" in {r.code for r in change.confidence.reasons}
            assert change.review.unit in ("row", "cell")
        else:
            assert change.confidence.final_color is was  # không đụng ô không xanh
    # Phá thử: bỏ trần thì ô xanh quay lại, chứng tỏ chính trần đã hạ chúng.
    assert Color.GREEN in {c.confidence.final_color for c in _fill(tmp_path, max_color="green").changes}


def test_cap_is_the_default_argument():
    import inspect

    assert inspect.signature(fill_from_excel).parameters["max_color"].default == "yellow"


def test_cap_function_alone():
    assert Color.GREEN not in {c.confidence.final_color for c in _cap_green(_green_batch()).changes}


# ---- Không lưu tên người, giá trị người bệnh (acceptance_lat1 2.4, 2.11)


def _source_with_people(tmp_path: Path) -> tuple[Path, list[str]]:
    wb = openpyxl.load_workbook(X01 / "source.xlsx")
    ws = wb.active
    name_col, id_col = ws.max_column + 1, ws.max_column + 2
    ws.cell(1, name_col, "Họ và tên")
    ws.cell(1, id_col, "Mã BN")
    values = []
    for row in range(2, ws.max_row + 1):
        name, code = f"Nguyễn Giả Bệnh {row:04d}", f"BN-GIA-{row:06d}"
        ws.cell(row, name_col, name)
        ws.cell(row, id_col, code)
        values += [name, code]
    path = tmp_path / "nguon.xlsx"
    wb.save(path)
    return path, values


def test_batch_has_no_person_name_and_no_patient_value_anywhere(tmp_path):
    source, patient_values = _source_with_people(tmp_path)
    batch = _fill(tmp_path, source=source)
    dumped = batch.model_dump_json()
    leaked = [v for v in (*PERSON_NAMES, *patient_values) if v in dumped]
    assert not leaked, leaked[:3]
    reasons = {n.reason for n in batch.notices}
    assert {"PERSON_COLUMN_LEFT_BLANK", "PHI_COLUMN_SKIPPED"} <= reasons
    assert len(batch.changes) == 180  # cột người bệnh không cản việc điền phần số


def test_person_names_really_are_in_the_source():
    """Nếu nguồn không có tên giả thì test trên không chứng minh gì."""
    text = json.dumps([[c.value for c in r] for r in openpyxl.load_workbook(X01 / "source.xlsx").active.iter_rows()], default=str, ensure_ascii=False)
    assert all(name in text for name in PERSON_NAMES)


# ---- Lỗi: luôn có .code, .detail, .message; không có tên file hay nội dung nguồn


def _assert_clean(exc: FillError, *secrets: str):
    assert exc.code and exc.message and "{" not in exc.message and "Biểu mẫu không bị đổi." in exc.message
    for secret in secrets:
        assert secret not in exc.detail and secret not in exc.message, secret


@pytest.mark.parametrize("name,code", [("GIA-X19_co_mat_khau.xlsx", "FILE_PASSWORD"), ("GIA-X19_hong.xlsx", "FILE_CORRUPT")])
def test_file_errors_have_a_user_message_and_no_file_name(tmp_path, name, code):
    source = tmp_path / "KQ_Tran_Gia_Hai.xlsx"
    shutil.copyfile(TEMPLATES / name, source)
    with pytest.raises(FillError) as exc:
        fill_from_excel(X01 / "template.xlsx", source, spec_id="GIA-QC-01", source_name="KQ_Tran_Gia_Hai.xlsx")
    assert exc.value.code == code
    _assert_clean(exc.value, "Tran_Gia_Hai", str(tmp_path))


def test_missing_fields_names_only_the_form_side(tmp_path):
    wb = openpyxl.load_workbook(X01 / "source.xlsx")
    head = [c.value for c in wb.active[1]]
    wb.active.cell(1, head.index("Kết quả") + 1, "Trị số Giả Riêng")
    path = tmp_path / "nguon.xlsx"
    wb.save(path)
    with pytest.raises(FillError) as exc:
        _fill(tmp_path, source=path)
    assert exc.value.code == "MISSING_FIELDS"
    assert exc.value.detail == "Chưa ghép được: Kết quả"
    assert "Kết quả" in exc.value.message
    _assert_clean(exc.value, "Trị số Giả Riêng")


@pytest.mark.parametrize("spec_id", ["KHONG-CO", "../config/llm", "GIA-QC-01/../GIA-QC-01", ""])
def test_unknown_or_unsafe_spec_id_is_refused(tmp_path, spec_id):
    with pytest.raises(FillError) as exc:
        fill_from_excel(X01 / "template.xlsx", X01 / "source.xlsx", spec_id=spec_id)
    assert exc.value.code == "SPEC_INVALID"
    _assert_clean(exc.value)


def test_unexpected_error_does_not_leak_its_text(monkeypatch, tmp_path):
    def boom(path):
        raise RuntimeError("Nguyễn Giả Lộ ở ô D7")

    monkeypatch.setattr(fill_module, "read_source", boom)
    with pytest.raises(FillError) as exc:
        _fill(tmp_path)
    assert exc.value.code == "FILL_FAILED" and exc.value.detail == "RuntimeError"
    _assert_clean(exc.value, "Nguyễn", "D7")


def test_batch_over_the_limit_is_refused_not_cut(monkeypatch, tmp_path):
    monkeypatch.setattr(fill_module, "MAX_CHANGES", 100)
    with pytest.raises(FillError) as exc:
        _fill(tmp_path)
    assert exc.value.code == "BATCH_TOO_LARGE" and "100" in exc.value.message
    _assert_clean(exc.value)


def test_template_error_detail_has_no_path(tmp_path):
    bad = tmp_path / "Mau_Le_Gia_Tu.xlsx"
    shutil.copyfile(TEMPLATES / "GIA-X19_hong.xlsx", bad)
    with pytest.raises(FillError) as exc:
        fill_from_excel(bad, X01 / "source.xlsx", spec_id="GIA-QC-01")
    assert exc.value.code == "TEMPLATE_UNREADABLE"
    _assert_clean(exc.value, "Le_Gia_Tu", str(tmp_path))


def test_every_code_fill_can_raise_has_a_full_message():
    codes = ["FILE_PASSWORD", "FILE_CORRUPT", "FILE_LEGACY_XLS", "FILE_TOO_LARGE", "NO_MATCHING_TABLE", "MISSING_FIELDS", "SPEC_INVALID", "SPEC_MISMATCH", "TEMPLATE_UNREADABLE", "BATCH_TOO_LARGE", "FILL_FAILED", "FEATURE_NOT_RELEASED"]
    texts = json.loads(fill_module.TEXTS.read_text(encoding="utf-8"))
    for code in codes:
        assert texts["reason_texts"].get(code), code
        message = user_message(code, {"cột": "Kết quả", "giới hạn": "tối đa 5.000 ô"})
        assert "{" not in message and "Biểu mẫu không bị đổi." in message, code
    assert texts["reason_texts"].get("MAP_GREEN_DISABLED")


def test_spec_is_the_packaged_one():
    assert fill_module._spec("GIA-QC-01") == load_spec(fill_module.SPECS / "GIA-QC-01.yaml")


# ---- Nguồn giả nghiệm thu lát 1 (T-AI-20)

LAT1 = ROOT / "fixtures" / "templates" / "nguon_lat1"
LAT1_FILES = sorted(LAT1.glob("*.xlsx"))


def test_lat1_sources_exist_and_match_the_generator(tmp_path, monkeypatch):
    """File trong repo phải đúng là thứ script sinh ra, để PO/BA và test cùng nói về một bộ."""
    spec = importlib.util.spec_from_file_location("gen_lat1", ROOT / "fixtures" / "eval" / "generate_lat1_sources.py")
    gen = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(gen)
    monkeypatch.setattr(gen, "OUT", tmp_path)
    gen.main()

    def cells(path):
        return [[str(v) for v in row] for row in openpyxl.load_workbook(path).active.iter_rows(values_only=True)]

    made = sorted(p.name for p in tmp_path.glob("*.xlsx"))
    assert made == sorted(p.name for p in LAT1_FILES) and len(made) == 9
    assert all(cells(tmp_path / name) == cells(LAT1 / name) for name in made)


@pytest.mark.parametrize("path", LAT1_FILES, ids=lambda p: p.stem)
def test_lat1_source_never_green_and_never_carries_people(tmp_path, path):
    batch = _fill(tmp_path, template=TEMPLATES / "GIA-QC-01_mau.xlsx", source=path)
    assert Color.GREEN not in {c.confidence.final_color for c in batch.changes}
    dumped = batch.model_dump_json()
    for fake in ("Giả Một", "Giả Hai", "Giả Ba", "Bệnh Nhân Giả", "BN-GIA-", *PERSON_NAMES):
        assert fake not in dumped, (path.name, fake)


# ---- Một ô một đề xuất (D10, X04)


def test_two_runs_on_one_day_give_one_red_change_with_both_readings(tmp_path):
    batch = _fill(tmp_path, template=TEMPLATES / "GIA-QC-01_mau.xlsx", source=LAT1 / "X04_mot_ngay_hai_lan_chay.xlsx")
    dup = [c for c in batch.changes if "SOURCE_DUPLICATE_RESULT" in {r.code for r in c.confidence.reasons}]
    assert len(dup) == 1 and len(batch.changes) == 180
    change = dup[0]
    assert change.confidence.final_color is Color.RED and change.value is None and change.review.unit == "cell"
    normalized = [e.normalized for e in change.evidence]
    assert len(normalized) == 2 and len(set(normalized)) == 2 and all(n and "," not in n for n in normalized)
    # Số chuẩn hóa của từng lần đọc khớp chữ trong nguồn, để máy chủ ghi đúng lần đọc người dùng chọn.
    assert all(float(e.normalized) == float(str(e.read_text).replace(",", ".")) for e in change.evidence)
    others = [c for c in batch.changes if c is not change]
    assert all(e.normalized is None for c in others for e in c.evidence)


def _all_qc_sources():
    generated = sorted((ROOT / "fixtures" / "eval" / "generated" / "xlsx").glob("X*/source.xlsx"))
    return [(p.parent / "template.xlsx", p) for p in generated if (p.parent / "template.xlsx").exists()] + [
        (TEMPLATES / "GIA-QC-01_mau.xlsx", p) for p in LAT1_FILES
    ]


def test_no_batch_has_two_changes_for_the_same_cell(tmp_path):
    seen = 0
    for template, source in _all_qc_sources():
        try:
            batch = fill_from_excel(template, source, spec_id="GIA-QC-01")
        except FillError:
            continue
        cells = Counter(c.target.a1_at_write for c in batch.changes)
        assert not [a1 for a1, n in cells.items() if n > 1], source.parent.name
        seen += 1
    assert seen >= 20  # đủ nhiều ca để có nghĩa, gồm X04 của cả hai bộ


def test_unknown_test_name_notice_gives_just_the_name_for_the_lat1_sentence(tmp_path):
    """X02 lát 1: tên xét nghiệm chưa khai thì các ô của tên đó không có đề xuất; `detail` chỉ là tên, để điền {tên}."""
    batch = _fill(tmp_path, template=TEMPLATES / "GIA-QC-01_mau.xlsx", source=LAT1 / "X02_doi_ten_cot.xlsx")
    notice = next(n for n in batch.notices if n.reason == "MAP_TEST_UNKNOWN")
    assert notice.detail == "Đường máu (giả)"
    assert not [c for c in batch.changes if c.target.a1_at_write[0] in "BC"]  # cột GLU L1, L2
    message = user_message("MAP_TEST_UNKNOWN", {"tên": notice.detail})
    assert message.startswith("File nguồn có tên xét nghiệm app chưa biết: Đường máu (giả).") and "{" not in message
    texts = json.loads(fill_module.TEXTS.read_text(encoding="utf-8"))
    assert texts["actions_lat1"]["MAP_COLUMN_INFERRED"].startswith("So số với file nguồn")


def _source_with_codes(tmp_path: Path, replace: dict[int, str]) -> Path:
    """Nguồn X01, thay giá trị cột Mã XN ở các dòng cho trước. [MẪU GIẢ]"""
    wb = openpyxl.load_workbook(X01 / "source.xlsx")
    ws = wb.active
    col = [c.value for c in ws[1]].index("Mã XN") + 1
    for row, value in replace.items():
        ws.cell(row, col, value)
    path = tmp_path / "nguon_ma_la.xlsx"
    wb.save(path)
    return path


def test_many_unknown_names_show_three_then_a_count(tmp_path):
    names = {r: f"Xét nghiệm lạ {chr(65 + r % 26)}{chr(65 + r // 26)}" for r in range(2, 42)}
    batch = _fill(tmp_path, source=_source_with_codes(tmp_path, names))
    detail = next(n.detail for n in batch.notices if n.reason == "MAP_TEST_UNKNOWN")
    assert detail.count(",") == 2 and detail.endswith(f"và {len(set(names.values())) - 3} tên khác")


def test_person_like_or_patient_like_value_in_the_test_column_is_never_shown(tmp_path):
    secrets = {2: "Nguyễn Giả Lẫn", 3: "BN-GIA-000123", 4: "Mã BN 20260901001"}
    batch = _fill(tmp_path, source=_source_with_codes(tmp_path, secrets))
    detail = next(n.detail for n in batch.notices if n.reason == "MAP_TEST_UNKNOWN")
    assert detail == "3 tên, app không hiện ở đây"
    dumped = batch.model_dump_json()
    assert not [s for s in secrets.values() if s in dumped]


def test_some_names_shown_and_some_hidden(tmp_path):
    batch = _fill(tmp_path, source=_source_with_codes(tmp_path, {2: "Đường máu (giả)", 3: "Nguyễn Giả Lẫn", 4: "BN-GIA-000777"}))
    detail = next(n.detail for n in batch.notices if n.reason == "MAP_TEST_UNKNOWN")
    assert detail == "Đường máu (giả) và 2 tên khác không hiện ở đây"


# ---- Chỗ trống của câu thông báo lát 1 (nghiệm thu K2, K3, N2, N3; lô 0.4.1, Notice.items)


def _render(code: str, notice) -> str:
    """Như giao diện (Dev, đã chốt 11/11): điền chỗ trống từ items, count, detail; ngày 08/09, tháng 10/2026."""
    def day(iso: str) -> str:
        return f"{iso[8:10]}/{iso[5:7]}"

    values = {"n": str(notice.count), "tên": notice.detail}
    if code == "SOURCE_MISSING_DAYS":
        values["ngày"] = ", ".join(day(i) for i in notice.items)
    elif code == "SOURCE_WRONG_MONTH":
        values["tháng nguồn"] = ", ".join(f"{i[5:7]}/{i[:4]}" for i in notice.items)
    elif code in ("PHI_COLUMN_SKIPPED", "PERSON_COLUMN_LEFT_BLANK"):
        values["cột"] = ", ".join(notice.items)
    texts = json.loads(fill_module.TEXTS.read_text(encoding="utf-8"))
    sentence = texts["texts_lat1"].get(code) or texts["reason_texts"][code]
    for key, value in values.items():
        sentence = sentence.replace("{" + key + "}", value)
    assert "{" not in sentence, (code, sentence)  # không chỗ trống nào thiếu giá trị
    return sentence


def _notice(tmp_path, name: str, code: str):
    batch = _fill(tmp_path, template=TEMPLATES / "GIA-QC-01_mau.xlsx", source=LAT1 / name)
    return batch, next(n for n in batch.notices if n.reason == code)


def test_missing_days_notice_gives_each_day_once_as_iso(tmp_path):
    batch, notice = _notice(tmp_path, "X03_thieu_3_ngay.xlsx", "SOURCE_MISSING_DAYS")
    assert notice.count == len(notice.items) == 3
    assert notice.items == sorted(set(notice.items))
    assert all(len(i) == 10 and i[4] == i[7] == "-" and i.startswith("2026-09-") for i in notice.items)
    filled = {p.value for c in batch.changes for p in c.target.row_key}
    assert not filled & set(notice.items)  # đúng là các ngày không có đề xuất
    sentence = _render("SOURCE_MISSING_DAYS", notice)
    assert sentence.startswith("Nguồn thiếu: ") and all(f"{i[8:10]}/09" in sentence for i in notice.items)


def test_wrong_month_notice_gives_count_and_source_month(tmp_path):
    _, notice = _notice(tmp_path, "X13_sai_thang_10.xlsx", "SOURCE_WRONG_MONTH")
    assert notice.items == ["2026-10"] and notice.count > 0
    assert _render("SOURCE_WRONG_MONTH", notice) == (
        f"File nguồn có {notice.count} dòng thuộc tháng 10/2026, không thuộc tháng của biểu mẫu nên app chưa điền."
    )


def test_patient_column_notice_names_the_columns_and_promises_nothing_about_sending(tmp_path):
    batch, notice = _notice(tmp_path, "X17_cot_nguoi_benh_gia.xlsx", "PHI_COLUMN_SKIPPED")
    assert notice.items == ["Họ tên người bệnh", "Mã BN"] and notice.count == 2
    sentence = _render("PHI_COLUMN_SKIPPED", notice)
    assert sentence == "File nguồn có cột giống thông tin người bệnh: Họ tên người bệnh, Mã BN. App không dùng các cột này."
    assert user_message("PHI_COLUMN_SKIPPED", {"cột": ", ".join(notice.items)}).endswith("Kiểm tra file có nên dùng không.")


def test_person_column_notice_names_the_column_once_with_no_repeated_action(tmp_path):
    _, notice = _notice(tmp_path, "nguoi_chay_ten_gia.xlsx", "PERSON_COLUMN_LEFT_BLANK")
    assert notice.items == ["Người chạy"] and notice.count == 1  # trước đây "Người chạy, Người chạy"
    assert _render("PERSON_COLUMN_LEFT_BLANK", notice) == "Cột Người chạy để anh/chị tự viết, AI không điền."
    message = user_message("PERSON_COLUMN_LEFT_BLANK", {"cột": "Người chạy"})
    assert message == "Cột Người chạy để anh/chị tự viết, AI không điền. Biểu mẫu không bị đổi."  # không còn "Tự điền."


@pytest.mark.parametrize("path", LAT1_FILES, ids=lambda p: p.stem)
def test_no_lat1_sentence_promises_not_sending_and_items_are_never_cell_values(tmp_path, path):
    batch = _fill(tmp_path, template=TEMPLATES / "GIA-QC-01_mau.xlsx", source=path)
    texts = json.loads(fill_module.TEXTS.read_text(encoding="utf-8"))
    for notice in batch.notices:
        assert "gửi ra ngoài" not in notice.detail
        if notice.reason in texts["reason_texts"]:
            assert "gửi ra ngoài" not in (texts["texts_lat1"].get(notice.reason) or texts["reason_texts"][notice.reason])
        assert len(notice.items) == len(set(notice.items)) <= 200 and all(len(i) <= 120 for i in notice.items)
        for fake in ("Giả", "BN-GIA-"):
            assert not any(fake in i for i in notice.items), (path.name, notice.reason)


def test_no_sentence_in_any_table_promises_not_sending_out():
    """PM 11/11 (US-1.3, Q94): app không hứa điều đội chưa chứng minh được, ở mọi lát, mọi bảng câu."""
    texts = json.loads(fill_module.TEXTS.read_text(encoding="utf-8"))
    tables = ("reason_texts", "texts_lat1", "actions", "actions_lat1", "ui_texts")
    found = [(t, k) for t in tables for k, s in texts[t].items() if "gửi ra ngoài" in s.lower()]
    assert not found, found
