"""infer_spec (Q53): các tính chất không được vỡ. Bố cục là tự nghĩ, không chứng minh được với mẫu thật."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path

import pytest

from vmai import Color, TemplateSpec, infer_spec, plan_headers, profile_from_guess, propose, read_form, read_source
from vmai.source_xlsx import SourceFacts

ROOT = Path(__file__).resolve().parents[2]
LAYOUTS = ROOT / "fixtures" / "eval" / "generated" / "layouts"
pytestmark = pytest.mark.skipif(not LAYOUTS.exists(), reason="Chưa tạo: chạy fixtures/eval/generate_layouts.py")


def _run_layouts():
    spec = importlib.util.spec_from_file_location("run_layouts", ROOT / "vmai" / "eval" / "run_layouts.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _facts(name: str):
    directory = LAYOUTS / name
    return directory, read_source(directory / "template.xlsx"), read_source(directory / "source.xlsx")


def test_every_layout_never_green_before_confirm_and_never_wrong_silently():
    module = _run_layouts()
    results = [module.run_case(d) for d in sorted(LAYOUTS.iterdir()) if d.is_dir()]
    assert len(results) >= 15
    for r in results:
        assert r["colors"].get("green", 0) == 0, r["case"]
        assert r["outcome"] not in ("đoán sai không báo", "hỏi 1 câu rồi vẫn sai"), r["case"]
        assert not r["forbidden_leak"], r["case"]
        assert not r["confirmed_wrong_green"], r["case"]


def test_questions_mean_no_spec_so_no_proposal_can_be_made():
    for name in ("L12_ten_khac_han", "L13_khoa_trung", "L18_muc_bang_chu", "L20_hai_cot_ngay", "L25_nguon_them_chieu"):
        _, template, source = _facts(name)
        guess = infer_spec(template, source)
        assert guess.status == "ask" and guess.spec is None and guess.questions, name


def test_weak_level_match_asks_and_can_be_accepted():
    _, template, source = _facts("L02_mot_dong_tieu_de")
    first = infer_spec(template, source)
    assert [q.code for q in first.questions] == ["LEVEL_NAME_DIFFERS"] and first.spec is None
    second = infer_spec(template, source, frozenset({"LEVEL_NAME_DIFFERS"}))
    assert second.status == "ok" and second.spec is not None


def test_confirmation_is_required_and_only_then_q17_decides_green():
    directory, template, source = _facts("L01_hai_dong_tieu_de")
    guess = infer_spec(template, source)
    form = read_form(directory / "template.xlsx", guess.spec)
    plan = plan_headers(form, source)
    assert all(c.confidence.final_color is Color.YELLOW for c in propose(form, source, plan).changes)
    with pytest.raises(ValueError):
        profile_from_guess(guess, form.fingerprint, plan.headers, confirmed_by="", confirmed_at="t", profile_id="p")

    profile = profile_from_guess(guess, form.fingerprint, plan.headers, confirmed_by="tester.gia", confirmed_at="t", profile_id="p")
    confirmed = propose(form, source, plan_headers(form, source, profile))
    assert all(c.confidence.final_color is Color.GREEN for c in confirmed.changes)
    assert confirmed.changes[0].mapping.profile_id == "p"


def test_confirmed_guess_without_limits_stays_yellow_by_q17():
    directory, template, source = _facts("L03_ba_dong_tieu_de")
    guess = infer_spec(template, source)
    form = read_form(directory / "template.xlsx", guess.spec)
    plan = plan_headers(form, source)
    profile = profile_from_guess(guess, form.fingerprint, plan.headers, confirmed_by="t", confirmed_at="t", profile_id="p")
    batch = propose(form, source, plan_headers(form, source, profile))
    assert batch.changes and all(c.confidence.final_color is Color.YELLOW for c in batch.changes)


def test_guess_and_facts_cross_as_json():
    directory, template, source = _facts("L05_danh_sach_vat_tu")
    t2 = SourceFacts.model_validate_json(template.model_dump_json())
    s2 = SourceFacts.model_validate_json(source.model_dump_json())
    guess = infer_spec(t2, s2)
    assert TemplateSpec.model_validate_json(guess.spec.model_dump_json()) == guess.spec
    assert guess.spec.origin == "inferred"
    assert guess.computed_columns == ("F",)
    json.dumps(guess.model_dump(mode="json"))


def test_person_and_note_columns_are_left_for_the_user():
    _, template, source = _facts("L15_cot_nguoi")
    guess = infer_spec(template, source, frozenset({"LEVEL_NAME_DIFFERS"}))
    assert {"F", "G", "H"} <= set(guess.manual_columns)
    filled = {c.column for c in guess.spec.table.value_columns}
    assert not filled & {"F", "G", "H"}
    assert "Người chạy" not in json.dumps(guess.spec.source_fields, default=str, ensure_ascii=False)


def test_spec_id_is_stable_between_months_so_a_profile_stays_valid(tmp_path):
    import openpyxl

    directory, template, source = _facts("L01_hai_dong_tieu_de")
    first = infer_spec(template, source)
    wb = openpyxl.load_workbook(directory / "source.xlsx")
    ws = wb.active
    for row in ws.iter_rows(min_row=2, min_col=4, max_col=4):
        row[0].value = 5.1
    wb.save(tmp_path / "other.xlsx")
    assert infer_spec(template, read_source(tmp_path / "other.xlsx")).spec.spec_id == first.spec.spec_id
