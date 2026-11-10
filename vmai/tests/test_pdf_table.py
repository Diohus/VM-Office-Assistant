"""Đọc bảng PDF có lớp chữ, sau cờ riêng (Q98). Bẫy tạo bằng fixtures/eval/generate_pdf_traps.py. PDF giả, không chứng minh gì với PDF thật."""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from vmai import Color, infer_spec, plan_headers, propose, read_form, read_source
from vmai.confidence import ReadFacts, SourceKind, grade_read
from vmai.errors import FeatureNotReleased, SourceFileError
from vmai.pdf_table import read_pdf_table

GENERATED = Path(__file__).resolve().parents[2] / "fixtures" / "eval" / "generated"
TRAPS = GENERATED / "pdf_traps"
P01 = GENERATED / "pdf" / "P01"
UNTRUSTED = [
    "T1_chu_an_tr3",
    "T2_chu_trang",
    "T3_chu_rat_nho",
    "T4_hai_chu_de_nhau",
    "T5_chu_ngoai_trang",
    "T6_anh_de_chu",
    "T7_trang_quet",
    "T8_chu_an_trong_form",
    "T9_o_trang_che_chu",
]


def _batch(pdf: Path, template: Path):
    source = read_pdf_table(pdf, experimental_pdf_sources=True)
    guess = infer_spec(read_source(template), source, frozenset({"LEVEL_NAME_DIFFERS"}), experimental_pdf_sources=True)
    assert guess.spec is not None, [q.code for q in guess.questions]
    form = read_form(template, guess.spec)
    return source, form, propose(form, source, plan_headers(form, source), experimental_pdf_sources=True)


def test_flag_off_refuses_before_opening_the_file(tmp_path):
    with pytest.raises(FeatureNotReleased):
        read_pdf_table(tmp_path / "khong_co.pdf")


def test_flag_is_off_by_default_and_separate_from_the_image_flag():
    with pytest.raises(FeatureNotReleased):
        read_pdf_table(P01 / "source.pdf")
    source = read_pdf_table(P01 / "source.pdf", experimental_pdf_sources=True)
    template = read_source(P01 / "template.xlsx")
    with pytest.raises(FeatureNotReleased):
        infer_spec(template, source)
    with pytest.raises(FeatureNotReleased):
        infer_spec(template, source, experimental_image_sources=True)
    guess = infer_spec(template, source, frozenset({"LEVEL_NAME_DIFFERS"}), experimental_pdf_sources=True)
    form = read_form(P01 / "template.xlsx", guess.spec)
    with pytest.raises(FeatureNotReleased):
        propose(form, source, plan_headers(form, source), experimental_image_sources=True)


def test_the_pdf_flag_does_not_open_images():
    from vmai.table_read import source_from_tokens

    tokens = [{"text": t, "bbox": [50 + i * 100, 100 + j * 30, 90 + i * 100, 120 + j * 30]} for j, row in enumerate([("A", "B"), ("1", "2"), ("3", "4"), ("5", "6")]) for i, t in enumerate(row)]
    source = source_from_tokens("a.jpg", "x", [(1, tokens, tokens)], kind="image", first_reader="ocr:a@1", second_reader="ocr:b@1")
    with pytest.raises(FeatureNotReleased):
        infer_spec(read_source(P01 / "template.xlsx"), source, experimental_pdf_sources=True)


def test_own_fixture_reads_every_cell_at_most_yellow():
    _, _, batch = _batch(P01 / "source.pdf", P01 / "template.xlsx")
    gold = {c["a1"]: c["value"] for c in json.loads((P01 / "answer.json").read_text(encoding="utf-8"))["cells"]}
    assert len(batch.changes) == len(gold) == 180
    assert {c.confidence.final_color for c in batch.changes} == {Color.YELLOW}
    assert all(float(c.value.normalized) == float(gold[c.target.a1_at_write]) for c in batch.changes)


def test_never_green_until_the_text_layer_is_compared_with_ocr():
    assert grade_read(ReadFacts(SourceKind.PDF_TEXT)).color is Color.YELLOW
    assert grade_read(ReadFacts(SourceKind.PDF_TEXT, pdf_text_verified=True)).color is Color.GREEN
    assert "pdf_text_verified=True" not in "".join(p.read_text(encoding="utf-8") for p in Path(__file__).resolve().parents[1].joinpath("src", "vmai").glob("*.py"))


def test_confirmed_profile_still_does_not_make_pdf_green():
    from vmai.confidence import ColumnMatch, RowMatch, combine, grade_map

    assert combine(grade_read(ReadFacts(SourceKind.PDF_TEXT)), grade_map(ColumnMatch.PROFILE, RowMatch.KEY_EXACT)).final_color is Color.YELLOW


@pytest.mark.parametrize("name", UNTRUSTED)
def test_text_layer_that_differs_from_what_is_shown_is_refused(name):
    with pytest.raises(SourceFileError) as exc:
        read_pdf_table(TRAPS / f"{name}.pdf", experimental_pdf_sources=True)
    assert exc.value.code == "PDF_TEXT_LAYER_UNTRUSTED"


def test_ordinary_pdf_is_not_refused():
    _, _, batch = _batch(TRAPS / "S0_binh_thuong.pdf", TRAPS / "template.xlsx")
    gold = json.loads((TRAPS / "gold.json").read_text(encoding="utf-8"))
    assert len(batch.changes) == len(gold)
    assert all(float(c.value.normalized) == float(gold[c.target.a1_at_write]) for c in batch.changes)


def test_not_a_pdf_is_reported(tmp_path):
    bad = tmp_path / "x.pdf"
    bad.write_bytes(b"khong phai pdf")
    with pytest.raises(SourceFileError) as exc:
        read_pdf_table(bad, experimental_pdf_sources=True)
    assert exc.value.code == "FILE_CORRUPT"
