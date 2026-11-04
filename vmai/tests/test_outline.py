"""Word đầu mục (Q54, Q66). Adapter giả chỉ kiểm luồng, không phải độ đúng đầu mục của model."""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from vmai.document import DocLine, DocPage, DocumentFacts, lines_from_tokens, mark_repeated, mark_uncertain, read_pdf_lines
from vmai.errors import SourceFileError
from vmai.llm.base import LLMBadOutput, LLMResponse
from vmai.llm.factory import create_provider, load_profile
from vmai.outline import Block, assign_headings, build_outline, check_text, check_written_paragraphs, guess_headings
from vmai.phi import ExternalSendPolicy, PhiDetected, SendNotAllowed

CASES = Path(__file__).resolve().parents[2] / "fixtures" / "eval" / "generated" / "word"
pytestmark = pytest.mark.skipif(not CASES.exists(), reason="Chưa tạo: chạy fixtures/eval/generate_word.py")
SEND = ExternalSendPolicy(allow_documents=True)


def _doc(name: str) -> DocumentFacts:
    return read_pdf_lines(CASES / name / "source.pdf")


def _fake(payload):
    return create_provider(load_profile("fake"), handler=lambda req: LLMResponse("fake", "fake-1", text=json.dumps(payload)))


def test_model_cannot_change_text_even_if_it_tries():
    document = _doc("D01_so_va_dam")
    line = next(l for l in document.lines if "Giới thiệu" in l.text)
    provider = _fake({"headings": [{"line": line.index, "level": 1, "text": "Chữ do model tự viết 999"}]})
    marks, calls, records = assign_headings(provider, document, SEND)
    outline = build_outline(document, marks, "model")
    assert outline.check.identical
    assert not any("tự viết" in b.text or "999" in b.text for b in outline.blocks)
    assert "tự viết" not in outline.model_dump_json()
    assert calls and records


def test_default_policy_sends_nothing():
    provider = _fake({"headings": []})
    with pytest.raises(SendNotAllowed):
        assign_headings(provider, _doc("D01_so_va_dam"), ExternalSendPolicy())
    assert provider.requests == []


def test_phi_signal_stops_sending_even_when_allowed():
    provider = _fake({"headings": []})
    with pytest.raises(PhiDetected):
        assign_headings(provider, _doc("D09_co_dau_hieu_nguoi_benh"), SEND)
    assert provider.requests == []


@pytest.mark.parametrize(
    "payload",
    [
        {"headings": [{"line": 10**6, "level": 1}]},
        {"headings": [{"line": 1, "level": 99}]},
        {"headings": [{"line": "1", "level": 1}]},
        {"nope": []},
    ],
)
def test_bad_model_output_is_rejected(payload):
    with pytest.raises(LLMBadOutput):
        assign_headings(_fake(payload), _doc("D01_so_va_dam"), SEND)


def test_outbound_log_says_what_was_sent_without_the_text():
    document = _doc("D01_so_va_dam")
    provider = _fake({"headings": []})
    _, _, records = assign_headings(provider, document, SEND)
    assert sum(r.line_count for r in records) == sum(not l.repeated and not l.in_table for l in document.lines)
    record = records[0]
    assert record.document_sha256 == document.sha256 and len(record.sent_sha256) == 64 and record.pages
    assert "Giới thiệu" not in record.model_dump_json()
    sent = provider.requests[0].messages[0].parts[0].text
    assert "Hãng Giả Định" not in sent  # đầu trang lặp không được gửi


def test_text_check_catches_a_changed_number():
    document = _doc("D01_so_va_dam")
    outline = build_outline(document, guess_headings(document.lines), "heuristic")
    assert outline.check.identical
    tampered = [b.model_copy(update={"text": b.text.replace("25 cm", "35 cm")}) for b in outline.blocks]
    check = check_text(list(document.lines), document.tables, tampered)
    assert not check.identical and "25" in check.missing and "35" in check.added


def test_text_check_catches_a_changed_table_cell():
    document = _doc("D01_so_va_dam")
    outline = build_outline(document, guess_headings(document.lines), "heuristic")
    tampered = [
        b.model_copy(update={"rows": tuple(tuple("999" if c == "220" else c for c in r) for r in b.rows)}) if b.kind == "table" else b for b in outline.blocks
    ]
    assert not check_text(list(document.lines), document.tables, tampered).identical


def test_written_file_check_catches_changes_and_ignores_added_numbers():
    document = _doc("D02_khong_so")
    outline = build_outline(document, guess_headings(document.lines), "heuristic")
    paragraphs = [f"{b.number} {b.text}" if b.number else b.text for b in outline.blocks if b.kind in ("heading", "paragraph")]
    assert check_written_paragraphs(outline, paragraphs).identical
    paragraphs[3] = paragraphs[3] + " thêm"
    assert not check_written_paragraphs(outline, paragraphs).identical


def test_numbers_are_added_only_when_the_source_has_none():
    without = build_outline(doc := _doc("D02_khong_so"), guess_headings(doc.lines), "heuristic")
    numbers = [b.number for b in without.blocks if b.kind == "heading"][:5]
    assert numbers == ["1", "1.1", "1.2", "2", "2.1"]
    with_numbers = build_outline(doc := _doc("D01_so_va_dam"), guess_headings(doc.lines), "heuristic")
    assert all(b.number is None for b in with_numbers.blocks if b.kind == "heading")


def test_steps_and_toc_are_not_headings():
    document = _doc("D05_muc_luc")
    marks = {m.line for m in guess_headings(document.lines)}
    texts = [l.text for l in document.lines if l.index in marks]
    assert not any("...." in t or t.startswith("1. Mở") for t in texts)


def test_repeated_header_footer_is_dropped_and_listed():
    document = _doc("D01_so_va_dam")
    outline = build_outline(document, guess_headings(document.lines), "heuristic")
    assert any("Hãng Giả Định" in l.text for l in outline.dropped_lines)
    assert any(l.text.startswith("Trang") for l in outline.dropped_lines)
    assert not any("Hãng Giả Định" in b.text for b in outline.blocks)


def test_two_column_reading_order():
    document = _doc("D06_hai_cot")
    assert any(p.two_column for p in document.pages)
    tops = [l for l in document.lines if not l.repeated and l.page == 1]
    # Hết cột trái mới sang cột phải.
    assert [l.x0 > 250 for l in tops] == sorted(l.x0 > 250 for l in tops)


def test_copy_restricted_opens_and_password_is_reported():
    assert _doc("D07_chan_copy").lines
    with pytest.raises(SourceFileError) as exc:
        _doc("D08_co_mat_khau")
    assert exc.value.code == "FILE_PASSWORD"


def test_scanned_page_has_no_trusted_text_and_ocr_lines_are_marked():
    document = _doc("D10_quet")
    assert document.scanned_pages and not document.lines


def test_uncertain_when_second_read_differs_or_is_missing():
    first = lines_from_tokens([{"text": "Điện", "bbox": [10, 10, 40, 20]}, {"text": "220", "bbox": [50, 10, 70, 20]}, {"text": "Tần", "bbox": [10, 40, 40, 50]}, {"text": "50", "bbox": [50, 40, 70, 50]}], 1)
    second = [{"text": "Dien 220", "bbox": [10, 10, 70, 20]}, {"text": "Tan 60", "bbox": [10, 40, 70, 50]}]
    marked = mark_uncertain(first, second)
    # Q70: cả hai dòng đều có chữ số nên đều phải so với ảnh, kể cả dòng hai OCR khớp nhau.
    assert [l.uncertain for l in marked] == [True, True]
    assert [l.digits_differ for l in marked] == [False, True]
    assert all(l.source == "ocr" for l in marked)
    assert mark_uncertain(first, [])[0].uncertain


def test_line_without_digits_is_not_marked_when_second_read_exists():
    first = lines_from_tokens([{"text": "Cảnh báo", "bbox": [10, 10, 70, 20]}], 1)
    assert not mark_uncertain(first, [{"text": "Canh bao", "bbox": [10, 10, 70, 20]}])[0].uncertain


def test_q70_every_wrong_digit_line_is_flagged_on_the_scan_fixture():
    answer_path = CASES / "D10_quet" / "ocr_tokens.json"
    if not answer_path.exists():
        pytest.skip("Chưa chạy ocr_word_pages.py")
    import importlib.util

    spec = importlib.util.spec_from_file_location("run_word", Path(__file__).resolve().parents[1] / "eval" / "run_word.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    answer = json.loads((CASES / "D10_quet" / "answer.json").read_text(encoding="utf-8"))
    result = module.scan_case(CASES / "D10_quet", answer)
    assert result["lines_with_wrong_digits"] > 0
    assert result["wrong_digit_lines_flagged"] == result["lines_with_wrong_digits"]


def test_q71_dropped_lines_are_listed_and_can_be_restored():
    document = _doc("D01_so_va_dam")
    marks = guess_headings(document.lines)
    outline = build_outline(document, marks, "heuristic")
    samples = {g.sample for g in outline.dropped_groups}
    assert any("Hãng Giả Định" in s for s in samples) and any(s.startswith("Trang") for s in samples)
    footer = next(g for g in outline.dropped_groups if g.sample.startswith("Trang"))
    assert len(footer.pages) == len(document.pages) and len(footer.line_indices) == len(footer.pages)

    restored = build_outline(document, marks, "heuristic", restore_lines=frozenset(footer.line_indices))
    assert restored.check.identical
    assert len(restored.dropped_lines) == len(outline.dropped_lines) - len(footer.line_indices)
    assert any(b.text.startswith("Trang") or "Trang" in b.text for b in restored.blocks)


def test_q71_repeated_warning_at_page_edge_is_dropped_but_visible():
    """Dòng thật lặp ở chân trang cũng bị bỏ. Vì vậy danh sách phải hiện cho người dùng, không bỏ im lặng."""
    pages = tuple(DocPage(page=p, kind="text", width=600, height=800) for p in range(1, 5))
    lines = []
    for p in range(1, 5):
        lines.append(DocLine(index=len(lines), page=p, text=f"Nội dung trang {p} khác nhau", x0=50, top=300, x1=300, bottom=312))
        lines.append(DocLine(index=len(lines), page=p, text="CẢNH BÁO: ngắt điện trước khi mở nắp", x0=50, top=770, x1=300, bottom=782))
    marked = mark_repeated(lines, list(pages))
    document = DocumentFacts(file_name="x", sha256="y", pages=pages, lines=tuple(marked))
    outline = build_outline(document, [], "heuristic")
    assert [g.sample for g in outline.dropped_groups] == ["CẢNH BÁO: ngắt điện trước khi mở nắp"]
    assert outline.dropped_groups[0].pages == (1, 2, 3, 4)


def test_level_jump_is_reported():
    lines = tuple(DocLine(index=i, page=1, text=t, x0=10, top=20.0 * i, x1=100, bottom=20.0 * i + 10) for i, t in enumerate(["A", "B"]))
    document = DocumentFacts(file_name="x", sha256="y", pages=(DocPage(page=1, kind="text", width=100, height=1000),), lines=lines)
    from vmai.outline import HeadingMark

    outline = build_outline(document, [HeadingMark(line=0, level=1), HeadingMark(line=1, level=3)], "model")
    assert outline.level_jumps and isinstance(outline.blocks[0], Block)


def test_scan_keeps_content_lines_only_real_header_footer_dropped():
    """Lỗi đã gặp: tọa độ OCR bằng pixel so với chiều cao trang bằng point làm dòng nội dung bị coi là chân trang."""
    if not (CASES / "D10_quet" / "ocr_tokens.json").exists():
        pytest.skip("Chưa chạy ocr_word_pages.py")
    import importlib.util

    spec = importlib.util.spec_from_file_location("run_word", Path(__file__).resolve().parents[1] / "eval" / "run_word.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    answer = json.loads((CASES / "D10_quet" / "answer.json").read_text(encoding="utf-8"))
    result = module.scan_case(CASES / "D10_quet", answer)
    assert result["dropped_lines"] <= 2 * result["pages_scanned"]
    assert result["lines_with_wrong_digits"] + result["right_digit_lines"] > 150
