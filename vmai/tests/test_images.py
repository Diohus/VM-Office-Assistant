"""Phần đọc ảnh, sau cờ (Q81, Q84). Dùng token giả tự tạo nên không cần OCR. Không chứng minh chất lượng trên ảnh thật."""
from __future__ import annotations

import subprocess

import openpyxl
import pytest

from vmai import Color, infer_spec, plan_headers, propose, read_form, read_source
from vmai.errors import SourceFileError
from vmai.infer import repair_header
from vmai.table_read import orientation_score, orientation_sure, source_from_tokens, tokens_to_grid

DAYS = [f"{d:02d}/09/2026" for d in range(1, 6)]
HEADERS = ["Ngày", "GLU M1", "GLU M2"]


def _template(path, headers=("Glucose", "Mức 1", "Mức 2")):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "[MẪU GIẢ] BIỂU MẪU"
    ws.merge_cells("A3:A4")
    ws["A3"] = "Ngày"
    ws.merge_cells("B3:C3")
    ws["B3"] = headers[0]
    ws["B4"], ws["C4"] = headers[1], headers[2]
    import datetime as dt

    for i in range(5):
        ws[f"A{5 + i}"] = dt.date(2026, 9, 1 + i)
    wb.save(path)
    return path


def _tokens(headers=HEADERS, values=None, second=False):
    xs = [50, 300, 500]
    out = [{"text": h, "bbox": [xs[i], 100, xs[i] + 90, 120]} for i, h in enumerate(headers)]
    for r, day in enumerate(DAYS):
        y = 130 + r * 30
        out.append({"text": day, "bbox": [xs[0], y, xs[0] + 110, y + 20]})
        row = (values or {}).get(r, ("5,10", "16,20"))
        for c, text in enumerate(row, start=1):
            out.append({"text": text, "bbox": [xs[c], y, xs[c] + 50, y + 20]})
    return out


def _source(first, second):
    return source_from_tokens("a.jpg", "x", [(1, first, second)], kind="image", first_reader="ocr:a@1", second_reader="ocr:b@1")


def _run(tmp_path, first, second, accepted=frozenset({"LEVEL_NAME_DIFFERS"})):
    template = _template(tmp_path / "t.xlsx")
    source = _source(first, second)
    guess = infer_spec(read_source(template), source, accepted, experimental_image_sources=True)
    if guess.spec is None:
        return guess, None
    form = read_form(template, guess.spec)
    return guess, propose(form, source, plan_headers(form, source), experimental_image_sources=True)


def test_clean_image_flow_is_yellow_never_green(tmp_path):
    guess, batch = _run(tmp_path, _tokens(), _tokens())
    assert guess.status == "ok"
    values = {c.target.a1_at_write: c.value.normalized for c in batch.changes}
    assert values["B5"] == "5.10" and values["C9"] == "16.20"
    assert all(c.confidence.final_color is Color.YELLOW for c in batch.changes)
    assert {e.reader for e in batch.changes[0].evidence} == {"ocr:a@1", "ocr:b@1"}
    assert batch.changes[0].evidence[0].locator.type == "image_region"


def test_reads_that_disagree_are_red(tmp_path):
    _, batch = _run(tmp_path, _tokens(), _tokens(values={0: ("5,70", "16,20")}))
    by_a1 = {c.target.a1_at_write: c for c in batch.changes}
    assert by_a1["B5"].confidence.final_color is Color.RED
    assert by_a1["C5"].confidence.final_color is Color.YELLOW


@pytest.mark.parametrize("text,fixed", [("ALTMI", "alt m1"), ("ALT M1", "alt m1"), ("ALTM1", "alt m1"), ("GLU Ml", "glu m1"), ("GLUM2", "glu m2"), ("Ngày", "ngay"), ("ALTM7", "alt m7")])
def test_repair_header(text, fixed):
    assert repair_header(text) == fixed


def test_q84_repaired_header_still_never_green_even_when_read_agrees(tmp_path):
    first = _tokens(["Ngày", "GLUMI", "GLU M2"])
    second = _tokens(["Ngày", "GLUM1", "GLUM2"])
    guess, batch = _run(tmp_path, first, second)
    assert guess.status == "ok"
    assert batch and all(c.confidence.final_color is not Color.GREEN for c in batch.changes)
    assert {c.target.a1_at_write[0] for c in batch.changes} == {"B", "C"}


def test_q84_unmapped_column_is_asked_alone_others_still_proposed(tmp_path):
    guess, batch = _run(tmp_path, _tokens(["Ngày", "GLU M1", "XXQQ"]), _tokens(["Ngày", "GLU M1", "XXQQ"]))
    assert guess.status == "partial" and guess.unmapped_columns == ("Glucose / Mức 2",)
    assert {c.target.a1_at_write[0] for c in batch.changes} == {"B"}


def test_q84_header_misread_as_other_level_with_second_read_disagreeing_is_not_mapped(tmp_path):
    """Lần đọc đầu đọc M2 thành M1 ở cột thứ hai: hai cột cùng là M1, không được ghép bừa."""
    guess, batch = _run(tmp_path, _tokens(["Ngày", "GLU M1", "GLU M1"]), _tokens())
    assert "B" not in {c.column for c in (guess.spec.table.value_columns if guess.spec else ())} or batch is None or all(c.confidence.final_color is not Color.GREEN for c in batch.changes)
    if batch is not None:
        for change in batch.changes:
            # Không có số nào của cột M2 được đưa vào cột M1.
            assert not (change.target.a1_at_write.startswith("B") and change.value.normalized == "16.20")


def test_q84_header_swapped_consistently_by_one_engine_stays_out(tmp_path):
    """Một engine đọc đảo M1, M2, engine kia đọc đúng: hai tiêu đề không khớp nên không ghép theo chúng."""
    guess, batch = _run(tmp_path, _tokens(["Ngày", "GLU M2", "GLU M1"]), _tokens())
    assert batch is None or not batch.changes


def test_missing_second_read_is_an_error_not_a_silent_fallback():
    with pytest.raises(SourceFileError) as exc:
        source_from_tokens("a.jpg", "x", [(1, _tokens(), None)], kind="image", first_reader="ocr:a@1", second_reader="ocr:b@1")
    assert exc.value.code == "OCR_SECOND_READ_MISSING"
    with pytest.raises(SourceFileError):
        source_from_tokens("a.jpg", "x", [(1, _tokens(), _tokens())], kind="image", first_reader="ocr:a@1", second_reader=None)


def test_no_table_is_reported():
    with pytest.raises(SourceFileError) as exc:
        _source([{"text": "Thông báo họp", "bbox": [10, 10, 200, 30]}], [])
    assert exc.value.code == "NO_MATCHING_TABLE"


def test_engine_failure_is_reported(monkeypatch, tmp_path):
    pytest.importorskip("cv2")
    from vmai.ocr import TesseractOcr

    fake = tmp_path / "tess.exe"
    fake.write_text("x")
    monkeypatch.delenv("VMAI_TESSERACT_DOCKER", raising=False)
    monkeypatch.setenv("VMAI_TESSERACT", str(fake))

    def boom(*args, **kwargs):
        raise subprocess.CalledProcessError(1, "tesseract", stderr=b"noi dung tai lieu")

    monkeypatch.setattr(subprocess, "run", boom)
    with pytest.raises(SourceFileError) as exc:
        TesseractOcr()
    assert exc.value.code == "OCR_ENGINE_FAILED" and "noi dung" not in exc.value.detail


def test_unconfigured_engine_is_reported(monkeypatch):
    pytest.importorskip("cv2")
    from vmai.ocr import TesseractOcr

    monkeypatch.delenv("VMAI_TESSERACT_DOCKER", raising=False)
    monkeypatch.delenv("VMAI_TESSERACT", raising=False)
    with pytest.raises(SourceFileError) as exc:
        TesseractOcr()
    assert exc.value.code == "OCR_NOT_INSTALLED"


def test_orientation_prefers_header_above_numbers():
    upright = _tokens()
    flipped = [{**t, "bbox": [t["bbox"][0], 400 - t["bbox"][3], t["bbox"][2], 400 - t["bbox"][1]]} for t in upright]
    assert orientation_score(upright) > orientation_score(flipped)


def test_orientation_prefers_readable_decimals_over_more_cells():
    """Ảnh ngược: Tesseract ra nhiều ô rác có chữ, tiêu đề rác vẫn đứng trên. Số thập phân đọc được mới quyết định."""
    upright = _tokens()
    junk = [{**t, "text": "Bug" if not t["text"][0].isdigit() else "80'91"} for t in upright]
    assert orientation_score(upright) > orientation_score(junk)
    assert orientation_sure(upright) and not orientation_sure(junk)


def test_integer_only_table_is_never_sure_so_all_turns_are_tried():
    assert not orientation_sure([{"text": "118", "bbox": [i * 50, 0, i * 50 + 30, 20]} for i in range(30)])


def test_page_title_cut_into_cells_by_one_engine_is_not_taken_as_header(tmp_path):
    title = [{"text": t, "bbox": [x, 60, x + 90, 80]} for t, x in (("PHIEU", 50), ("KET", 300), ("QUA", 500))]
    guess, batch = _run(tmp_path, title + _tokens(), _tokens())
    assert guess.status == "ok"
    assert {c.target.a1_at_write for c in batch.changes} >= {"B5", "C9"}


def test_header_directly_above_data_is_kept_and_all_rows_stay():
    source = _source(_tokens(), _tokens())
    assert len(source.sheets[0].rows) == 1 + len(DAYS)


def test_half_shadowed_paper_is_not_inverted():
    np = pytest.importorskip("numpy")
    pytest.importorskip("cv2")
    from vmai.ocr import _upright_gray

    paper = np.full((200, 200), 240, dtype=np.uint8)
    paper[:, 100:] = 90
    paper[::20, 10:190] = 10
    assert _upright_gray(paper)[5, 5] == 240


def test_dim_photo_is_not_inverted_but_dark_screen_is():
    np = pytest.importorskip("numpy")
    pytest.importorskip("cv2")
    from vmai.ocr import _upright_gray

    dim_paper = np.full((100, 100), 80, dtype=np.uint8)
    dim_paper[40:45, 10:90] = 15
    dark_screen = np.full((100, 100), 30, dtype=np.uint8)
    dark_screen[40:45, 10:90] = 220
    assert _upright_gray(dim_paper)[0, 0] == 80
    assert _upright_gray(dark_screen)[0, 0] == 225


def test_row_with_text_outside_columns_is_left_out_of_grid():
    grid, lines = tokens_to_grid(_tokens() + [{"text": "Trang chú thích rải rác dài", "bbox": [120, 300, 280, 320]}])
    assert len(lines) == 3 and all(any(c is not None for c in row) for row in grid)


def _text_page(angle: float = 0.0):
    """Trang chữ không kẻ bảng, vẽ bằng OpenCV để không phụ thuộc font trên máy."""
    import cv2
    import numpy as np

    page = np.full((1200, 1600, 3), 255, dtype=np.uint8)
    for r in range(22):
        for c in range(5):
            cv2.putText(page, f"{r * 7 + c}.{r}", (100 + c * 280, 90 + r * 48), cv2.FONT_HERSHEY_SIMPLEX, 0.9, (0, 0, 0), 2, cv2.LINE_AA)
    if angle:
        page = cv2.warpAffine(page, cv2.getRotationMatrix2D((800, 600), angle, 1.0), (1600, 1200), borderValue=(255, 255, 255))
    return page


def _save(tmp_path, array, name="p.png"):
    from PIL import Image

    path = tmp_path / name
    Image.fromarray(array).save(path)
    return path


@pytest.mark.parametrize("angle", [5.0, -5.0, 2.0])
def test_table_without_lines_is_straightened_by_its_text_rows(tmp_path, angle):
    from vmai.ocr import prepare_image

    page = prepare_image(_save(tmp_path, _text_page(angle)))
    assert page.skew_degrees == pytest.approx(-angle, abs=0.5)


def test_straight_table_without_lines_is_not_turned(tmp_path):
    from vmai.ocr import prepare_image

    assert prepare_image(_save(tmp_path, _text_page())).skew_degrees == 0.0


def test_blank_page_is_not_turned(tmp_path):
    import numpy as np

    from vmai.ocr import prepare_image

    assert prepare_image(_save(tmp_path, np.full((800, 1000, 3), 255, dtype=np.uint8))).skew_degrees == 0.0


def test_edge_ratio_is_recorded_but_does_not_decide_low_quality(tmp_path):
    """Q102: ngưỡng này khớp theo bộ đã thấy và sai trên bộ độc lập. Xem lại khi có ảnh thật."""
    import cv2

    from vmai.ocr import prepare_image

    sharp = _text_page()
    h, w = sharp.shape[:2]
    soft = cv2.resize(cv2.resize(sharp, (w // 3, h // 3), interpolation=cv2.INTER_AREA), (w, h), interpolation=cv2.INTER_LINEAR)
    page = prepare_image(_save(tmp_path, soft))
    assert page.edge_ratio < prepare_image(_save(tmp_path, sharp, "a.png")).edge_ratio
    assert not page.low_quality
